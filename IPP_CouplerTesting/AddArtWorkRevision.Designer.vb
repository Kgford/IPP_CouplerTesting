<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AddArtWorkRevision
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AddArtWorkRevision))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btrev = New System.Windows.Forms.Button()
        Me.txtQuadrant = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.txtPanel = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.txtArtwork = New System.Windows.Forms.TextBox()
        Me.lblArtwork = New System.Windows.Forms.TextBox()
        Me.txtLOT = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Black
        Me.Label1.Font = New System.Drawing.Font("Arial Black", 26.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(69, 71)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(356, 52)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Artwork Revision"
        '
        'PictureBox1
        '
        Me.PictureBox1.ErrorImage = CType(resources.GetObject("PictureBox1.ErrorImage"), System.Drawing.Image)
        Me.PictureBox1.Image = Global.IPP_CouplerTesting.My.Resources.Resources.ipplogo2_Automation1
        Me.PictureBox1.Location = New System.Drawing.Point(63, 12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(371, 64)
        Me.PictureBox1.TabIndex = 6
        Me.PictureBox1.TabStop = False
        '
        'btrev
        '
        Me.btrev.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btrev.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btrev.Location = New System.Drawing.Point(205, 176)
        Me.btrev.Name = "btrev"
        Me.btrev.Size = New System.Drawing.Size(75, 30)
        Me.btrev.TabIndex = 7
        Me.btrev.Text = "OK"
        Me.btrev.UseVisualStyleBackColor = True
        '
        'txtQuadrant
        '
        Me.txtQuadrant.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtQuadrant.Location = New System.Drawing.Point(301, 126)
        Me.txtQuadrant.Name = "txtQuadrant"
        Me.txtQuadrant.Size = New System.Drawing.Size(47, 24)
        Me.txtQuadrant.TabIndex = 69
        Me.txtQuadrant.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox4
        '
        Me.TextBox4.BackColor = System.Drawing.Color.Black
        Me.TextBox4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox4.ForeColor = System.Drawing.Color.White
        Me.TextBox4.Location = New System.Drawing.Point(230, 129)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(65, 17)
        Me.TextBox4.TabIndex = 68
        Me.TextBox4.Text = "Quadrant"
        Me.TextBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtPanel
        '
        Me.txtPanel.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPanel.Location = New System.Drawing.Point(177, 126)
        Me.txtPanel.Name = "txtPanel"
        Me.txtPanel.Size = New System.Drawing.Size(47, 24)
        Me.txtPanel.TabIndex = 67
        Me.txtPanel.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.Color.Black
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.ForeColor = System.Drawing.Color.White
        Me.TextBox2.Location = New System.Drawing.Point(131, 130)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(40, 17)
        Me.TextBox2.TabIndex = 66
        Me.TextBox2.Text = "Panel"
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtArtwork
        '
        Me.txtArtwork.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtArtwork.Location = New System.Drawing.Point(78, 126)
        Me.txtArtwork.Name = "txtArtwork"
        Me.txtArtwork.Size = New System.Drawing.Size(47, 24)
        Me.txtArtwork.TabIndex = 65
        Me.txtArtwork.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblArtwork
        '
        Me.lblArtwork.BackColor = System.Drawing.Color.Black
        Me.lblArtwork.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lblArtwork.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblArtwork.ForeColor = System.Drawing.Color.White
        Me.lblArtwork.Location = New System.Drawing.Point(12, 129)
        Me.lblArtwork.Name = "lblArtwork"
        Me.lblArtwork.Size = New System.Drawing.Size(60, 17)
        Me.lblArtwork.TabIndex = 64
        Me.lblArtwork.Text = "Artwork"
        Me.lblArtwork.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtLOT
        '
        Me.txtLOT.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLOT.Location = New System.Drawing.Point(400, 127)
        Me.txtLOT.Name = "txtLOT"
        Me.txtLOT.Size = New System.Drawing.Size(100, 24)
        Me.txtLOT.TabIndex = 71
        Me.txtLOT.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.Black
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.Color.White
        Me.TextBox1.Location = New System.Drawing.Point(354, 130)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(40, 17)
        Me.TextBox1.TabIndex = 72
        Me.TextBox1.Text = "LOT"
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'AddArtWorkRevision
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(504, 246)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.txtLOT)
        Me.Controls.Add(Me.txtQuadrant)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.txtPanel)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.txtArtwork)
        Me.Controls.Add(Me.lblArtwork)
        Me.Controls.Add(Me.btrev)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.Label1)
        Me.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AddArtWorkRevision"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "Who are you?"
        Me.TransparencyKey = System.Drawing.Color.DarkRed
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents btrev As System.Windows.Forms.Button
    Friend WithEvents txtQuadrant As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents txtPanel As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents txtArtwork As System.Windows.Forms.TextBox
    Friend WithEvents lblArtwork As System.Windows.Forms.TextBox
    Friend WithEvents txtLOT As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
