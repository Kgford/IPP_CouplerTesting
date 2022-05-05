<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MYMsg
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MYMsg))
        Me.btOK = New System.Windows.Forms.Button()
        Me.txttitle = New System.Windows.Forms.Label()
        Me.txtprompt = New System.Windows.Forms.Label()
        Me.btNO = New System.Windows.Forms.Button()
        Me.btCANCEL = New System.Windows.Forms.Button()
        Me.btYES = New System.Windows.Forms.Button()
        Me.btABORT = New System.Windows.Forms.Button()
        Me.btIGNORE = New System.Windows.Forms.Button()
        Me.btRETRY = New System.Windows.Forms.Button()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btOK
        '
        Me.btOK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btOK.Location = New System.Drawing.Point(44, 62)
        Me.btOK.Name = "btOK"
        Me.btOK.Size = New System.Drawing.Size(107, 53)
        Me.btOK.TabIndex = 7
        Me.btOK.Text = "OK"
        Me.btOK.UseVisualStyleBackColor = True
        '
        'txttitle
        '
        Me.txttitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txttitle.Location = New System.Drawing.Point(9, 7)
        Me.txttitle.Name = "txttitle"
        Me.txttitle.Size = New System.Drawing.Size(447, 43)
        Me.txttitle.TabIndex = 8
        Me.txttitle.Text = "Title"
        Me.txttitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtprompt
        '
        Me.txtprompt.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtprompt.Location = New System.Drawing.Point(9, 137)
        Me.txtprompt.Name = "txtprompt"
        Me.txtprompt.Size = New System.Drawing.Size(447, 66)
        Me.txtprompt.TabIndex = 9
        Me.txtprompt.Text = "Prompt"
        Me.txtprompt.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'btNO
        '
        Me.btNO.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btNO.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btNO.Location = New System.Drawing.Point(170, 62)
        Me.btNO.Name = "btNO"
        Me.btNO.Size = New System.Drawing.Size(107, 53)
        Me.btNO.TabIndex = 10
        Me.btNO.Text = "NO"
        Me.btNO.UseVisualStyleBackColor = True
        '
        'btCANCEL
        '
        Me.btCANCEL.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btCANCEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btCANCEL.Location = New System.Drawing.Point(296, 62)
        Me.btCANCEL.Name = "btCANCEL"
        Me.btCANCEL.Size = New System.Drawing.Size(107, 53)
        Me.btCANCEL.TabIndex = 11
        Me.btCANCEL.Text = "CANCEL"
        Me.btCANCEL.UseVisualStyleBackColor = True
        '
        'btYES
        '
        Me.btYES.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btYES.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btYES.Location = New System.Drawing.Point(44, 62)
        Me.btYES.Name = "btYES"
        Me.btYES.Size = New System.Drawing.Size(107, 53)
        Me.btYES.TabIndex = 12
        Me.btYES.Text = "YES"
        Me.btYES.UseVisualStyleBackColor = True
        '
        'btABORT
        '
        Me.btABORT.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btABORT.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btABORT.Location = New System.Drawing.Point(296, 62)
        Me.btABORT.Name = "btABORT"
        Me.btABORT.Size = New System.Drawing.Size(107, 53)
        Me.btABORT.TabIndex = 13
        Me.btABORT.Text = "ABORT"
        Me.btABORT.UseVisualStyleBackColor = True
        '
        'btIGNORE
        '
        Me.btIGNORE.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btIGNORE.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btIGNORE.Location = New System.Drawing.Point(296, 62)
        Me.btIGNORE.Name = "btIGNORE"
        Me.btIGNORE.Size = New System.Drawing.Size(107, 53)
        Me.btIGNORE.TabIndex = 14
        Me.btIGNORE.Text = "IGNORE"
        Me.btIGNORE.UseVisualStyleBackColor = True
        '
        'btRETRY
        '
        Me.btRETRY.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btRETRY.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btRETRY.Location = New System.Drawing.Point(170, 62)
        Me.btRETRY.Name = "btRETRY"
        Me.btRETRY.Size = New System.Drawing.Size(107, 53)
        Me.btRETRY.TabIndex = 15
        Me.btRETRY.Text = "RETRY"
        Me.btRETRY.UseVisualStyleBackColor = True
        '
        'PictureBox1
        '
        Me.PictureBox1.ErrorImage = CType(resources.GetObject("PictureBox1.ErrorImage"), System.Drawing.Image)
        Me.PictureBox1.Image = Global.IPP_CouplerTesting.My.Resources.Resources.ipplogo_burst400_jpg
        Me.PictureBox1.Location = New System.Drawing.Point(14, 206)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(442, 132)
        Me.PictureBox1.TabIndex = 16
        Me.PictureBox1.TabStop = False
        '
        'MYMsg
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.ClientSize = New System.Drawing.Size(481, 379)
        Me.ControlBox = False
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.btRETRY)
        Me.Controls.Add(Me.btIGNORE)
        Me.Controls.Add(Me.btABORT)
        Me.Controls.Add(Me.btYES)
        Me.Controls.Add(Me.btCANCEL)
        Me.Controls.Add(Me.btNO)
        Me.Controls.Add(Me.txtprompt)
        Me.Controls.Add(Me.txttitle)
        Me.Controls.Add(Me.btOK)
        Me.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MYMsg"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Text = "The choice is yours..... Choose wisely"
        Me.TransparencyKey = System.Drawing.Color.SandyBrown
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btOK As System.Windows.Forms.Button
    Friend WithEvents txttitle As System.Windows.Forms.Label
    Friend WithEvents txtprompt As System.Windows.Forms.Label
    Friend WithEvents btNO As System.Windows.Forms.Button
    Friend WithEvents btCANCEL As System.Windows.Forms.Button
    Friend WithEvents btYES As System.Windows.Forms.Button
    Friend WithEvents btABORT As System.Windows.Forms.Button
    Friend WithEvents btIGNORE As System.Windows.Forms.Button
    Friend WithEvents btRETRY As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
End Class
