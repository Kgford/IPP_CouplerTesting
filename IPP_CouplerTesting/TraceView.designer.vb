<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TraceView
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TraceView))
        Me.cmbExecute = New System.Windows.Forms.Button()
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
        Me.Trace1Sel = New System.Windows.Forms.ComboBox()
        Me.Trace1Select = New System.Windows.Forms.CheckBox()
        Me.Trace2Select = New System.Windows.Forms.CheckBox()
        Me.Trace2Sel = New System.Windows.Forms.ComboBox()
        Me.Trace3Select = New System.Windows.Forms.CheckBox()
        Me.Trace3Sel = New System.Windows.Forms.ComboBox()
        Me.Trace4Select = New System.Windows.Forms.CheckBox()
        Me.Trace4Sel = New System.Windows.Forms.ComboBox()
        Me.lblTestID1 = New System.Windows.Forms.Label()
        Me.lblSeialNumber1 = New System.Windows.Forms.Label()
        Me.lblWorkStation1 = New System.Windows.Forms.Label()
        Me.lblPoints1 = New System.Windows.Forms.Label()
        Me.lblDate1 = New System.Windows.Forms.Label()
        Me.txtTestID1 = New System.Windows.Forms.Label()
        Me.txtSerialNumber1 = New System.Windows.Forms.Label()
        Me.txtWorkStation1 = New System.Windows.Forms.Label()
        Me.txtPoints1 = New System.Windows.Forms.Label()
        Me.txtDate1 = New System.Windows.Forms.Label()
        Me.txtWorkStation2 = New System.Windows.Forms.Label()
        Me.lblTestD2 = New System.Windows.Forms.Label()
        Me.lblSeialNumber2 = New System.Windows.Forms.Label()
        Me.lblDate2 = New System.Windows.Forms.Label()
        Me.txtTestID2 = New System.Windows.Forms.Label()
        Me.txtPoints2 = New System.Windows.Forms.Label()
        Me.lblPoints2 = New System.Windows.Forms.Label()
        Me.lblWorkStation2 = New System.Windows.Forms.Label()
        Me.txtSerialNumber2 = New System.Windows.Forms.Label()
        Me.txtSerialNumber3 = New System.Windows.Forms.Label()
        Me.lblPoints3 = New System.Windows.Forms.Label()
        Me.lblWorkStation3 = New System.Windows.Forms.Label()
        Me.txtDate3 = New System.Windows.Forms.Label()
        Me.txtPoints3 = New System.Windows.Forms.Label()
        Me.txtWorkStation3 = New System.Windows.Forms.Label()
        Me.txtTestID3 = New System.Windows.Forms.Label()
        Me.lblDate3 = New System.Windows.Forms.Label()
        Me.lblSeialNumber3 = New System.Windows.Forms.Label()
        Me.lblTestID3 = New System.Windows.Forms.Label()
        Me.txtSerialNumber4 = New System.Windows.Forms.Label()
        Me.lblPoints4 = New System.Windows.Forms.Label()
        Me.lblWorkStation4 = New System.Windows.Forms.Label()
        Me.txtDate4 = New System.Windows.Forms.Label()
        Me.txtPoints4 = New System.Windows.Forms.Label()
        Me.txtWorkStation4 = New System.Windows.Forms.Label()
        Me.txtTestID4 = New System.Windows.Forms.Label()
        Me.lblDate4 = New System.Windows.Forms.Label()
        Me.lblSeialNumber4 = New System.Windows.Forms.Label()
        Me.lblTestID4 = New System.Windows.Forms.Label()
        Me.cmbSerialNumber1 = New System.Windows.Forms.ComboBox()
        Me.cmbSerialNumber2 = New System.Windows.Forms.ComboBox()
        Me.TraceGrid = New System.Windows.Forms.DataGridView()
        Me.ZG1 = New ZedGraph.ZedGraphControl()
        Me.txtDate2 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.TraceGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbExecute
        '
        Me.cmbExecute.BackColor = System.Drawing.SystemColors.ControlDark
        Me.cmbExecute.Location = New System.Drawing.Point(1087, 6)
        Me.cmbExecute.Name = "cmbExecute"
        Me.cmbExecute.Size = New System.Drawing.Size(125, 32)
        Me.cmbExecute.TabIndex = 1
        Me.cmbExecute.Text = "EXECUTE SEARCH"
        Me.cmbExecute.UseVisualStyleBackColor = False
        '
        'cmbPARM1
        '
        Me.cmbPARM1.AutoCompleteCustomSource.AddRange(New String() {"", "TestID", "JobNumber", "ModelNumber", "SerialNumber", "Operator"})
        Me.cmbPARM1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPARM1.FormattingEnabled = True
        Me.cmbPARM1.Items.AddRange(New Object() {"", "TraceID", "TestID", "Title", "JobNumber", "PartNumber", "SerialNumber", "WorkStation", "Operator"})
        Me.cmbPARM1.Location = New System.Drawing.Point(237, 45)
        Me.cmbPARM1.Name = "cmbPARM1"
        Me.cmbPARM1.Size = New System.Drawing.Size(118, 24)
        Me.cmbPARM1.TabIndex = 3
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Snow
        Me.Label1.Location = New System.Drawing.Point(9, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(199, 20)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Select * from Trace   where"
        '
        'cmbPARM2
        '
        Me.cmbPARM2.AutoCompleteCustomSource.AddRange(New String() {"", "TestID", "JobNumber", "ModelNumber", "SerialNumber", "Operator"})
        Me.cmbPARM2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPARM2.FormattingEnabled = True
        Me.cmbPARM2.Items.AddRange(New Object() {"", "TraceID", "TestID", "Title", "JobNumber", "PartNumber", "SerialNumber", "WorkStation", "Operator"})
        Me.cmbPARM2.Location = New System.Drawing.Point(561, 46)
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
        Me.E1.Location = New System.Drawing.Point(380, 49)
        Me.E1.Name = "E1"
        Me.E1.Size = New System.Drawing.Size(18, 20)
        Me.E1.TabIndex = 6
        Me.E1.Text = "="
        Me.E1.Visible = False
        '
        'txtPARM1
        '
        Me.txtPARM1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPARM1.Location = New System.Drawing.Point(416, 46)
        Me.txtPARM1.Name = "txtPARM1"
        Me.txtPARM1.Size = New System.Drawing.Size(88, 22)
        Me.txtPARM1.TabIndex = 7
        Me.txtPARM1.Visible = False
        '
        'A1
        '
        Me.A1.AutoSize = True
        Me.A1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.A1.ForeColor = System.Drawing.Color.Snow
        Me.A1.Location = New System.Drawing.Point(519, 49)
        Me.A1.Name = "A1"
        Me.A1.Size = New System.Drawing.Size(36, 20)
        Me.A1.TabIndex = 8
        Me.A1.Text = "and"
        Me.A1.Visible = False
        '
        'txtPARM2
        '
        Me.txtPARM2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPARM2.Location = New System.Drawing.Point(718, 49)
        Me.txtPARM2.Name = "txtPARM2"
        Me.txtPARM2.Size = New System.Drawing.Size(93, 22)
        Me.txtPARM2.TabIndex = 11
        Me.txtPARM2.Visible = False
        '
        'E2
        '
        Me.E2.AutoSize = True
        Me.E2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.E2.ForeColor = System.Drawing.Color.Snow
        Me.E2.Location = New System.Drawing.Point(680, 52)
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
        Me.Label5.Location = New System.Drawing.Point(525, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(272, 40)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "TRACE SEARCH"
        '
        'txtPARM3
        '
        Me.txtPARM3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPARM3.Location = New System.Drawing.Point(1049, 48)
        Me.txtPARM3.Name = "txtPARM3"
        Me.txtPARM3.Size = New System.Drawing.Size(93, 22)
        Me.txtPARM3.TabIndex = 52
        Me.txtPARM3.Visible = False
        '
        'E3
        '
        Me.E3.AutoSize = True
        Me.E3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.E3.ForeColor = System.Drawing.Color.Snow
        Me.E3.Location = New System.Drawing.Point(1003, 49)
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
        Me.A2.Location = New System.Drawing.Point(817, 48)
        Me.A2.Name = "A2"
        Me.A2.Size = New System.Drawing.Size(36, 20)
        Me.A2.TabIndex = 50
        Me.A2.Text = "and"
        Me.A2.Visible = False
        '
        'cmbPARM3
        '
        Me.cmbPARM3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbPARM3.FormattingEnabled = True
        Me.cmbPARM3.Items.AddRange(New Object() {"", "TraceID", "TestID", "Title", "JobNumber", "PartNumber", "SerialNumber", "WorkStation", "Operator"})
        Me.cmbPARM3.Location = New System.Drawing.Point(863, 48)
        Me.cmbPARM3.Name = "cmbPARM3"
        Me.cmbPARM3.Size = New System.Drawing.Size(119, 24)
        Me.cmbPARM3.TabIndex = 49
        Me.cmbPARM3.Visible = False
        '
        'Trace1Sel
        '
        Me.Trace1Sel.AutoCompleteCustomSource.AddRange(New String() {"", "TestID", "JobNumber", "ModelNumber", "SerialNumber", "Operator"})
        Me.Trace1Sel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Trace1Sel.FormattingEnabled = True
        Me.Trace1Sel.Items.AddRange(New Object() {""})
        Me.Trace1Sel.Location = New System.Drawing.Point(67, 88)
        Me.Trace1Sel.Name = "Trace1Sel"
        Me.Trace1Sel.Size = New System.Drawing.Size(190, 24)
        Me.Trace1Sel.TabIndex = 54
        Me.Trace1Sel.Visible = False
        '
        'Trace1Select
        '
        Me.Trace1Select.AutoSize = True
        Me.Trace1Select.ForeColor = System.Drawing.Color.Salmon
        Me.Trace1Select.Location = New System.Drawing.Point(1, 92)
        Me.Trace1Select.Name = "Trace1Select"
        Me.Trace1Select.Size = New System.Drawing.Size(63, 17)
        Me.Trace1Select.TabIndex = 55
        Me.Trace1Select.Text = "Trace 1"
        Me.Trace1Select.UseVisualStyleBackColor = True
        '
        'Trace2Select
        '
        Me.Trace2Select.AutoSize = True
        Me.Trace2Select.ForeColor = System.Drawing.Color.Salmon
        Me.Trace2Select.Location = New System.Drawing.Point(278, 92)
        Me.Trace2Select.Name = "Trace2Select"
        Me.Trace2Select.Size = New System.Drawing.Size(63, 17)
        Me.Trace2Select.TabIndex = 57
        Me.Trace2Select.Text = "Trace 2"
        Me.Trace2Select.UseVisualStyleBackColor = True
        '
        'Trace2Sel
        '
        Me.Trace2Sel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Trace2Sel.FormattingEnabled = True
        Me.Trace2Sel.Items.AddRange(New Object() {""})
        Me.Trace2Sel.Location = New System.Drawing.Point(347, 89)
        Me.Trace2Sel.Name = "Trace2Sel"
        Me.Trace2Sel.Size = New System.Drawing.Size(208, 24)
        Me.Trace2Sel.TabIndex = 56
        Me.Trace2Sel.Visible = False
        '
        'Trace3Select
        '
        Me.Trace3Select.AutoSize = True
        Me.Trace3Select.ForeColor = System.Drawing.Color.Salmon
        Me.Trace3Select.Location = New System.Drawing.Point(635, 89)
        Me.Trace3Select.Name = "Trace3Select"
        Me.Trace3Select.Size = New System.Drawing.Size(63, 17)
        Me.Trace3Select.TabIndex = 59
        Me.Trace3Select.Text = "Trace 3"
        Me.Trace3Select.UseVisualStyleBackColor = True
        '
        'Trace3Sel
        '
        Me.Trace3Sel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Trace3Sel.FormattingEnabled = True
        Me.Trace3Sel.Items.AddRange(New Object() {""})
        Me.Trace3Sel.Location = New System.Drawing.Point(704, 89)
        Me.Trace3Sel.Name = "Trace3Sel"
        Me.Trace3Sel.Size = New System.Drawing.Size(219, 24)
        Me.Trace3Sel.TabIndex = 58
        Me.Trace3Sel.Visible = False
        '
        'Trace4Select
        '
        Me.Trace4Select.AutoSize = True
        Me.Trace4Select.ForeColor = System.Drawing.Color.Salmon
        Me.Trace4Select.Location = New System.Drawing.Point(990, 89)
        Me.Trace4Select.Name = "Trace4Select"
        Me.Trace4Select.Size = New System.Drawing.Size(63, 17)
        Me.Trace4Select.TabIndex = 61
        Me.Trace4Select.Text = "Trace 4"
        Me.Trace4Select.UseVisualStyleBackColor = True
        '
        'Trace4Sel
        '
        Me.Trace4Sel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Trace4Sel.FormattingEnabled = True
        Me.Trace4Sel.Items.AddRange(New Object() {""})
        Me.Trace4Sel.Location = New System.Drawing.Point(1059, 85)
        Me.Trace4Sel.Name = "Trace4Sel"
        Me.Trace4Sel.Size = New System.Drawing.Size(194, 24)
        Me.Trace4Sel.TabIndex = 60
        Me.Trace4Sel.Visible = False
        '
        'lblTestID1
        '
        Me.lblTestID1.AutoSize = True
        Me.lblTestID1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTestID1.ForeColor = System.Drawing.Color.Snow
        Me.lblTestID1.Location = New System.Drawing.Point(48, 118)
        Me.lblTestID1.Name = "lblTestID1"
        Me.lblTestID1.Size = New System.Drawing.Size(39, 13)
        Me.lblTestID1.TabIndex = 62
        Me.lblTestID1.Text = "TestID"
        Me.lblTestID1.Visible = False
        '
        'lblSeialNumber1
        '
        Me.lblSeialNumber1.AutoSize = True
        Me.lblSeialNumber1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSeialNumber1.ForeColor = System.Drawing.Color.Snow
        Me.lblSeialNumber1.Location = New System.Drawing.Point(24, 131)
        Me.lblSeialNumber1.Name = "lblSeialNumber1"
        Me.lblSeialNumber1.Size = New System.Drawing.Size(70, 13)
        Me.lblSeialNumber1.TabIndex = 63
        Me.lblSeialNumber1.Text = "SerialNumber"
        Me.lblSeialNumber1.Visible = False
        '
        'lblWorkStation1
        '
        Me.lblWorkStation1.AutoSize = True
        Me.lblWorkStation1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWorkStation1.ForeColor = System.Drawing.Color.Snow
        Me.lblWorkStation1.Location = New System.Drawing.Point(28, 144)
        Me.lblWorkStation1.Name = "lblWorkStation1"
        Me.lblWorkStation1.Size = New System.Drawing.Size(66, 13)
        Me.lblWorkStation1.TabIndex = 64
        Me.lblWorkStation1.Text = "WorkStation"
        Me.lblWorkStation1.Visible = False
        '
        'lblPoints1
        '
        Me.lblPoints1.AutoSize = True
        Me.lblPoints1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPoints1.ForeColor = System.Drawing.Color.Snow
        Me.lblPoints1.Location = New System.Drawing.Point(58, 157)
        Me.lblPoints1.Name = "lblPoints1"
        Me.lblPoints1.Size = New System.Drawing.Size(36, 13)
        Me.lblPoints1.TabIndex = 66
        Me.lblPoints1.Text = "Points"
        Me.lblPoints1.Visible = False
        '
        'lblDate1
        '
        Me.lblDate1.AutoSize = True
        Me.lblDate1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate1.ForeColor = System.Drawing.Color.Snow
        Me.lblDate1.Location = New System.Drawing.Point(64, 170)
        Me.lblDate1.Name = "lblDate1"
        Me.lblDate1.Size = New System.Drawing.Size(30, 13)
        Me.lblDate1.TabIndex = 67
        Me.lblDate1.Text = "Date"
        Me.lblDate1.Visible = False
        '
        'txtTestID1
        '
        Me.txtTestID1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTestID1.ForeColor = System.Drawing.Color.Snow
        Me.txtTestID1.Location = New System.Drawing.Point(100, 118)
        Me.txtTestID1.Name = "txtTestID1"
        Me.txtTestID1.Size = New System.Drawing.Size(118, 13)
        Me.txtTestID1.TabIndex = 68
        Me.txtTestID1.Text = "TestID1"
        Me.txtTestID1.Visible = False
        '
        'txtSerialNumber1
        '
        Me.txtSerialNumber1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSerialNumber1.ForeColor = System.Drawing.Color.Snow
        Me.txtSerialNumber1.Location = New System.Drawing.Point(100, 131)
        Me.txtSerialNumber1.Name = "txtSerialNumber1"
        Me.txtSerialNumber1.Size = New System.Drawing.Size(118, 13)
        Me.txtSerialNumber1.TabIndex = 69
        Me.txtSerialNumber1.Text = "SerialNumber"
        Me.txtSerialNumber1.Visible = False
        '
        'txtWorkStation1
        '
        Me.txtWorkStation1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWorkStation1.ForeColor = System.Drawing.Color.Snow
        Me.txtWorkStation1.Location = New System.Drawing.Point(100, 144)
        Me.txtWorkStation1.Name = "txtWorkStation1"
        Me.txtWorkStation1.Size = New System.Drawing.Size(118, 13)
        Me.txtWorkStation1.TabIndex = 70
        Me.txtWorkStation1.Text = "WorkStation"
        Me.txtWorkStation1.Visible = False
        '
        'txtPoints1
        '
        Me.txtPoints1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPoints1.ForeColor = System.Drawing.Color.Snow
        Me.txtPoints1.Location = New System.Drawing.Point(100, 157)
        Me.txtPoints1.Name = "txtPoints1"
        Me.txtPoints1.Size = New System.Drawing.Size(118, 13)
        Me.txtPoints1.TabIndex = 72
        Me.txtPoints1.Text = "Points"
        Me.txtPoints1.Visible = False
        '
        'txtDate1
        '
        Me.txtDate1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDate1.ForeColor = System.Drawing.Color.Snow
        Me.txtDate1.Location = New System.Drawing.Point(100, 170)
        Me.txtDate1.Name = "txtDate1"
        Me.txtDate1.Size = New System.Drawing.Size(118, 13)
        Me.txtDate1.TabIndex = 73
        Me.txtDate1.Text = "Date"
        Me.txtDate1.Visible = False
        '
        'txtWorkStation2
        '
        Me.txtWorkStation2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWorkStation2.ForeColor = System.Drawing.Color.Snow
        Me.txtWorkStation2.Location = New System.Drawing.Point(361, 144)
        Me.txtWorkStation2.Name = "txtWorkStation2"
        Me.txtWorkStation2.Size = New System.Drawing.Size(118, 13)
        Me.txtWorkStation2.TabIndex = 80
        Me.txtWorkStation2.Text = "WorkStation"
        Me.txtWorkStation2.Visible = False
        '
        'lblTestD2
        '
        Me.lblTestD2.AutoSize = True
        Me.lblTestD2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTestD2.ForeColor = System.Drawing.Color.Snow
        Me.lblTestD2.Location = New System.Drawing.Point(309, 118)
        Me.lblTestD2.Name = "lblTestD2"
        Me.lblTestD2.Size = New System.Drawing.Size(39, 13)
        Me.lblTestD2.TabIndex = 75
        Me.lblTestD2.Text = "TestID"
        Me.lblTestD2.Visible = False
        '
        'lblSeialNumber2
        '
        Me.lblSeialNumber2.AutoSize = True
        Me.lblSeialNumber2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSeialNumber2.ForeColor = System.Drawing.Color.Snow
        Me.lblSeialNumber2.Location = New System.Drawing.Point(285, 131)
        Me.lblSeialNumber2.Name = "lblSeialNumber2"
        Me.lblSeialNumber2.Size = New System.Drawing.Size(70, 13)
        Me.lblSeialNumber2.TabIndex = 76
        Me.lblSeialNumber2.Text = "SerialNumber"
        Me.lblSeialNumber2.Visible = False
        '
        'lblDate2
        '
        Me.lblDate2.AutoSize = True
        Me.lblDate2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate2.ForeColor = System.Drawing.Color.Snow
        Me.lblDate2.Location = New System.Drawing.Point(325, 170)
        Me.lblDate2.Name = "lblDate2"
        Me.lblDate2.Size = New System.Drawing.Size(30, 13)
        Me.lblDate2.TabIndex = 78
        Me.lblDate2.Text = "Date"
        Me.lblDate2.Visible = False
        '
        'txtTestID2
        '
        Me.txtTestID2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTestID2.ForeColor = System.Drawing.Color.Snow
        Me.txtTestID2.Location = New System.Drawing.Point(361, 118)
        Me.txtTestID2.Name = "txtTestID2"
        Me.txtTestID2.Size = New System.Drawing.Size(118, 13)
        Me.txtTestID2.TabIndex = 79
        Me.txtTestID2.Text = "TestID1"
        Me.txtTestID2.Visible = False
        '
        'txtPoints2
        '
        Me.txtPoints2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPoints2.ForeColor = System.Drawing.Color.Snow
        Me.txtPoints2.Location = New System.Drawing.Point(361, 157)
        Me.txtPoints2.Name = "txtPoints2"
        Me.txtPoints2.Size = New System.Drawing.Size(118, 13)
        Me.txtPoints2.TabIndex = 82
        Me.txtPoints2.Text = "Points"
        Me.txtPoints2.Visible = False
        '
        'lblPoints2
        '
        Me.lblPoints2.AutoSize = True
        Me.lblPoints2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPoints2.ForeColor = System.Drawing.Color.Snow
        Me.lblPoints2.Location = New System.Drawing.Point(319, 157)
        Me.lblPoints2.Name = "lblPoints2"
        Me.lblPoints2.Size = New System.Drawing.Size(36, 13)
        Me.lblPoints2.TabIndex = 85
        Me.lblPoints2.Text = "Points"
        Me.lblPoints2.Visible = False
        '
        'lblWorkStation2
        '
        Me.lblWorkStation2.AutoSize = True
        Me.lblWorkStation2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWorkStation2.ForeColor = System.Drawing.Color.Snow
        Me.lblWorkStation2.Location = New System.Drawing.Point(289, 144)
        Me.lblWorkStation2.Name = "lblWorkStation2"
        Me.lblWorkStation2.Size = New System.Drawing.Size(66, 13)
        Me.lblWorkStation2.TabIndex = 84
        Me.lblWorkStation2.Text = "WorkStation"
        Me.lblWorkStation2.Visible = False
        '
        'txtSerialNumber2
        '
        Me.txtSerialNumber2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSerialNumber2.ForeColor = System.Drawing.Color.Snow
        Me.txtSerialNumber2.Location = New System.Drawing.Point(361, 131)
        Me.txtSerialNumber2.Name = "txtSerialNumber2"
        Me.txtSerialNumber2.Size = New System.Drawing.Size(118, 13)
        Me.txtSerialNumber2.TabIndex = 86
        Me.txtSerialNumber2.Text = "SerialNumber"
        Me.txtSerialNumber2.Visible = False
        '
        'txtSerialNumber3
        '
        Me.txtSerialNumber3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSerialNumber3.ForeColor = System.Drawing.Color.Snow
        Me.txtSerialNumber3.Location = New System.Drawing.Point(715, 131)
        Me.txtSerialNumber3.Name = "txtSerialNumber3"
        Me.txtSerialNumber3.Size = New System.Drawing.Size(118, 13)
        Me.txtSerialNumber3.TabIndex = 98
        Me.txtSerialNumber3.Text = "SerialNumber"
        Me.txtSerialNumber3.Visible = False
        '
        'lblPoints3
        '
        Me.lblPoints3.AutoSize = True
        Me.lblPoints3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPoints3.ForeColor = System.Drawing.Color.Snow
        Me.lblPoints3.Location = New System.Drawing.Point(673, 157)
        Me.lblPoints3.Name = "lblPoints3"
        Me.lblPoints3.Size = New System.Drawing.Size(36, 13)
        Me.lblPoints3.TabIndex = 97
        Me.lblPoints3.Text = "Points"
        Me.lblPoints3.Visible = False
        '
        'lblWorkStation3
        '
        Me.lblWorkStation3.AutoSize = True
        Me.lblWorkStation3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWorkStation3.ForeColor = System.Drawing.Color.Snow
        Me.lblWorkStation3.Location = New System.Drawing.Point(643, 144)
        Me.lblWorkStation3.Name = "lblWorkStation3"
        Me.lblWorkStation3.Size = New System.Drawing.Size(66, 13)
        Me.lblWorkStation3.TabIndex = 96
        Me.lblWorkStation3.Text = "WorkStation"
        Me.lblWorkStation3.Visible = False
        '
        'txtDate3
        '
        Me.txtDate3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDate3.ForeColor = System.Drawing.Color.Snow
        Me.txtDate3.Location = New System.Drawing.Point(715, 170)
        Me.txtDate3.Name = "txtDate3"
        Me.txtDate3.Size = New System.Drawing.Size(118, 13)
        Me.txtDate3.TabIndex = 95
        Me.txtDate3.Text = "Date"
        Me.txtDate3.Visible = False
        '
        'txtPoints3
        '
        Me.txtPoints3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPoints3.ForeColor = System.Drawing.Color.Snow
        Me.txtPoints3.Location = New System.Drawing.Point(715, 157)
        Me.txtPoints3.Name = "txtPoints3"
        Me.txtPoints3.Size = New System.Drawing.Size(118, 13)
        Me.txtPoints3.TabIndex = 94
        Me.txtPoints3.Text = "Points"
        Me.txtPoints3.Visible = False
        '
        'txtWorkStation3
        '
        Me.txtWorkStation3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWorkStation3.ForeColor = System.Drawing.Color.Snow
        Me.txtWorkStation3.Location = New System.Drawing.Point(715, 144)
        Me.txtWorkStation3.Name = "txtWorkStation3"
        Me.txtWorkStation3.Size = New System.Drawing.Size(118, 13)
        Me.txtWorkStation3.TabIndex = 92
        Me.txtWorkStation3.Text = "WorkStation"
        Me.txtWorkStation3.Visible = False
        '
        'txtTestID3
        '
        Me.txtTestID3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTestID3.ForeColor = System.Drawing.Color.Snow
        Me.txtTestID3.Location = New System.Drawing.Point(715, 118)
        Me.txtTestID3.Name = "txtTestID3"
        Me.txtTestID3.Size = New System.Drawing.Size(118, 13)
        Me.txtTestID3.TabIndex = 91
        Me.txtTestID3.Text = "TestID1"
        Me.txtTestID3.Visible = False
        '
        'lblDate3
        '
        Me.lblDate3.AutoSize = True
        Me.lblDate3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate3.ForeColor = System.Drawing.Color.Snow
        Me.lblDate3.Location = New System.Drawing.Point(679, 170)
        Me.lblDate3.Name = "lblDate3"
        Me.lblDate3.Size = New System.Drawing.Size(30, 13)
        Me.lblDate3.TabIndex = 90
        Me.lblDate3.Text = "Date"
        Me.lblDate3.Visible = False
        '
        'lblSeialNumber3
        '
        Me.lblSeialNumber3.AutoSize = True
        Me.lblSeialNumber3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSeialNumber3.ForeColor = System.Drawing.Color.Snow
        Me.lblSeialNumber3.Location = New System.Drawing.Point(639, 131)
        Me.lblSeialNumber3.Name = "lblSeialNumber3"
        Me.lblSeialNumber3.Size = New System.Drawing.Size(70, 13)
        Me.lblSeialNumber3.TabIndex = 88
        Me.lblSeialNumber3.Text = "SerialNumber"
        Me.lblSeialNumber3.Visible = False
        '
        'lblTestID3
        '
        Me.lblTestID3.AutoSize = True
        Me.lblTestID3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTestID3.ForeColor = System.Drawing.Color.Snow
        Me.lblTestID3.Location = New System.Drawing.Point(663, 118)
        Me.lblTestID3.Name = "lblTestID3"
        Me.lblTestID3.Size = New System.Drawing.Size(39, 13)
        Me.lblTestID3.TabIndex = 87
        Me.lblTestID3.Text = "TestID"
        Me.lblTestID3.Visible = False
        '
        'txtSerialNumber4
        '
        Me.txtSerialNumber4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSerialNumber4.ForeColor = System.Drawing.Color.Snow
        Me.txtSerialNumber4.Location = New System.Drawing.Point(1065, 131)
        Me.txtSerialNumber4.Name = "txtSerialNumber4"
        Me.txtSerialNumber4.Size = New System.Drawing.Size(118, 13)
        Me.txtSerialNumber4.TabIndex = 110
        Me.txtSerialNumber4.Text = "SerialNumber"
        Me.txtSerialNumber4.Visible = False
        '
        'lblPoints4
        '
        Me.lblPoints4.AutoSize = True
        Me.lblPoints4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPoints4.ForeColor = System.Drawing.Color.Snow
        Me.lblPoints4.Location = New System.Drawing.Point(1023, 157)
        Me.lblPoints4.Name = "lblPoints4"
        Me.lblPoints4.Size = New System.Drawing.Size(36, 13)
        Me.lblPoints4.TabIndex = 109
        Me.lblPoints4.Text = "Points"
        Me.lblPoints4.Visible = False
        '
        'lblWorkStation4
        '
        Me.lblWorkStation4.AutoSize = True
        Me.lblWorkStation4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblWorkStation4.ForeColor = System.Drawing.Color.Snow
        Me.lblWorkStation4.Location = New System.Drawing.Point(993, 144)
        Me.lblWorkStation4.Name = "lblWorkStation4"
        Me.lblWorkStation4.Size = New System.Drawing.Size(66, 13)
        Me.lblWorkStation4.TabIndex = 108
        Me.lblWorkStation4.Text = "WorkStation"
        Me.lblWorkStation4.Visible = False
        '
        'txtDate4
        '
        Me.txtDate4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDate4.ForeColor = System.Drawing.Color.Snow
        Me.txtDate4.Location = New System.Drawing.Point(1065, 170)
        Me.txtDate4.Name = "txtDate4"
        Me.txtDate4.Size = New System.Drawing.Size(118, 13)
        Me.txtDate4.TabIndex = 107
        Me.txtDate4.Text = "Date"
        Me.txtDate4.Visible = False
        '
        'txtPoints4
        '
        Me.txtPoints4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPoints4.ForeColor = System.Drawing.Color.Snow
        Me.txtPoints4.Location = New System.Drawing.Point(1065, 157)
        Me.txtPoints4.Name = "txtPoints4"
        Me.txtPoints4.Size = New System.Drawing.Size(118, 13)
        Me.txtPoints4.TabIndex = 106
        Me.txtPoints4.Text = "Points"
        Me.txtPoints4.Visible = False
        '
        'txtWorkStation4
        '
        Me.txtWorkStation4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtWorkStation4.ForeColor = System.Drawing.Color.Snow
        Me.txtWorkStation4.Location = New System.Drawing.Point(1065, 144)
        Me.txtWorkStation4.Name = "txtWorkStation4"
        Me.txtWorkStation4.Size = New System.Drawing.Size(118, 13)
        Me.txtWorkStation4.TabIndex = 104
        Me.txtWorkStation4.Text = "WorkStation"
        Me.txtWorkStation4.Visible = False
        '
        'txtTestID4
        '
        Me.txtTestID4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTestID4.ForeColor = System.Drawing.Color.Snow
        Me.txtTestID4.Location = New System.Drawing.Point(1065, 118)
        Me.txtTestID4.Name = "txtTestID4"
        Me.txtTestID4.Size = New System.Drawing.Size(118, 13)
        Me.txtTestID4.TabIndex = 103
        Me.txtTestID4.Text = "TestID1"
        Me.txtTestID4.Visible = False
        '
        'lblDate4
        '
        Me.lblDate4.AutoSize = True
        Me.lblDate4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDate4.ForeColor = System.Drawing.Color.Snow
        Me.lblDate4.Location = New System.Drawing.Point(1029, 170)
        Me.lblDate4.Name = "lblDate4"
        Me.lblDate4.Size = New System.Drawing.Size(30, 13)
        Me.lblDate4.TabIndex = 102
        Me.lblDate4.Text = "Date"
        Me.lblDate4.Visible = False
        '
        'lblSeialNumber4
        '
        Me.lblSeialNumber4.AutoSize = True
        Me.lblSeialNumber4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSeialNumber4.ForeColor = System.Drawing.Color.Snow
        Me.lblSeialNumber4.Location = New System.Drawing.Point(989, 131)
        Me.lblSeialNumber4.Name = "lblSeialNumber4"
        Me.lblSeialNumber4.Size = New System.Drawing.Size(70, 13)
        Me.lblSeialNumber4.TabIndex = 100
        Me.lblSeialNumber4.Text = "SerialNumber"
        Me.lblSeialNumber4.Visible = False
        '
        'lblTestID4
        '
        Me.lblTestID4.AutoSize = True
        Me.lblTestID4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTestID4.ForeColor = System.Drawing.Color.Snow
        Me.lblTestID4.Location = New System.Drawing.Point(1013, 118)
        Me.lblTestID4.Name = "lblTestID4"
        Me.lblTestID4.Size = New System.Drawing.Size(39, 13)
        Me.lblTestID4.TabIndex = 99
        Me.lblTestID4.Text = "TestID"
        Me.lblTestID4.Visible = False
        '
        'cmbSerialNumber1
        '
        Me.cmbSerialNumber1.AutoCompleteCustomSource.AddRange(New String() {"", "TestID", "JobNumber", "ModelNumber", "SerialNumber", "Operator"})
        Me.cmbSerialNumber1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSerialNumber1.FormattingEnabled = True
        Me.cmbSerialNumber1.Items.AddRange(New Object() {""})
        Me.cmbSerialNumber1.Location = New System.Drawing.Point(718, 46)
        Me.cmbSerialNumber1.Name = "cmbSerialNumber1"
        Me.cmbSerialNumber1.Size = New System.Drawing.Size(93, 24)
        Me.cmbSerialNumber1.TabIndex = 111
        Me.cmbSerialNumber1.Visible = False
        '
        'cmbSerialNumber2
        '
        Me.cmbSerialNumber2.AutoCompleteCustomSource.AddRange(New String() {"", "TestID", "JobNumber", "ModelNumber", "SerialNumber", "Operator"})
        Me.cmbSerialNumber2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSerialNumber2.FormattingEnabled = True
        Me.cmbSerialNumber2.Items.AddRange(New Object() {""})
        Me.cmbSerialNumber2.Location = New System.Drawing.Point(1049, 47)
        Me.cmbSerialNumber2.Name = "cmbSerialNumber2"
        Me.cmbSerialNumber2.Size = New System.Drawing.Size(93, 24)
        Me.cmbSerialNumber2.TabIndex = 112
        Me.cmbSerialNumber2.Visible = False
        '
        'TraceGrid
        '
        Me.TraceGrid.BackgroundColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.TraceGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.TraceGrid.Location = New System.Drawing.Point(1, 119)
        Me.TraceGrid.Name = "TraceGrid"
        Me.TraceGrid.Size = New System.Drawing.Size(1326, 466)
        Me.TraceGrid.TabIndex = 113
        '
        'ZG1
        '
        Me.ZG1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ZG1.EditButtons = System.Windows.Forms.MouseButtons.Left
        Me.ZG1.Location = New System.Drawing.Point(1, 135)
        Me.ZG1.Margin = New System.Windows.Forms.Padding(4)
        Me.ZG1.Name = "ZG1"
        Me.ZG1.PanModifierKeys = CType((System.Windows.Forms.Keys.Shift Or System.Windows.Forms.Keys.None), System.Windows.Forms.Keys)
        Me.ZG1.ScrollGrace = 0.0R
        Me.ZG1.ScrollMaxX = 0.0R
        Me.ZG1.ScrollMaxY = 0.0R
        Me.ZG1.ScrollMaxY2 = 0.0R
        Me.ZG1.ScrollMinX = 0.0R
        Me.ZG1.ScrollMinY = 0.0R
        Me.ZG1.ScrollMinY2 = 0.0R
        Me.ZG1.Size = New System.Drawing.Size(1326, 448)
        Me.ZG1.TabIndex = 114
        '
        'txtDate2
        '
        Me.txtDate2.BackColor = System.Drawing.Color.Black
        Me.txtDate2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDate2.ForeColor = System.Drawing.Color.White
        Me.txtDate2.Location = New System.Drawing.Point(361, 170)
        Me.txtDate2.Name = "txtDate2"
        Me.txtDate2.Size = New System.Drawing.Size(118, 13)
        Me.txtDate2.TabIndex = 115
        Me.txtDate2.Text = "Date"
        Me.txtDate2.Visible = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.ControlDark
        Me.Button1.Location = New System.Drawing.Point(1218, 6)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(94, 32)
        Me.Button1.TabIndex = 116
        Me.Button1.Text = "EXPORT DATA"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'TraceView
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.ClientSize = New System.Drawing.Size(1324, 587)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ZG1)
        Me.Controls.Add(Me.TraceGrid)
        Me.Controls.Add(Me.cmbSerialNumber2)
        Me.Controls.Add(Me.cmbSerialNumber1)
        Me.Controls.Add(Me.txtSerialNumber4)
        Me.Controls.Add(Me.lblPoints4)
        Me.Controls.Add(Me.lblWorkStation4)
        Me.Controls.Add(Me.txtDate4)
        Me.Controls.Add(Me.txtPoints4)
        Me.Controls.Add(Me.txtWorkStation4)
        Me.Controls.Add(Me.txtTestID4)
        Me.Controls.Add(Me.lblDate4)
        Me.Controls.Add(Me.lblSeialNumber4)
        Me.Controls.Add(Me.lblTestID4)
        Me.Controls.Add(Me.txtSerialNumber3)
        Me.Controls.Add(Me.lblPoints3)
        Me.Controls.Add(Me.lblWorkStation3)
        Me.Controls.Add(Me.txtDate3)
        Me.Controls.Add(Me.txtPoints3)
        Me.Controls.Add(Me.txtWorkStation3)
        Me.Controls.Add(Me.txtTestID3)
        Me.Controls.Add(Me.lblDate3)
        Me.Controls.Add(Me.lblSeialNumber3)
        Me.Controls.Add(Me.lblTestID3)
        Me.Controls.Add(Me.txtSerialNumber2)
        Me.Controls.Add(Me.lblPoints2)
        Me.Controls.Add(Me.lblWorkStation2)
        Me.Controls.Add(Me.txtPoints2)
        Me.Controls.Add(Me.txtWorkStation2)
        Me.Controls.Add(Me.txtTestID2)
        Me.Controls.Add(Me.lblDate2)
        Me.Controls.Add(Me.lblSeialNumber2)
        Me.Controls.Add(Me.lblTestD2)
        Me.Controls.Add(Me.txtDate1)
        Me.Controls.Add(Me.txtPoints1)
        Me.Controls.Add(Me.txtWorkStation1)
        Me.Controls.Add(Me.txtSerialNumber1)
        Me.Controls.Add(Me.txtTestID1)
        Me.Controls.Add(Me.lblDate1)
        Me.Controls.Add(Me.lblPoints1)
        Me.Controls.Add(Me.lblWorkStation1)
        Me.Controls.Add(Me.lblSeialNumber1)
        Me.Controls.Add(Me.lblTestID1)
        Me.Controls.Add(Me.Trace4Select)
        Me.Controls.Add(Me.Trace4Sel)
        Me.Controls.Add(Me.Trace3Select)
        Me.Controls.Add(Me.Trace3Sel)
        Me.Controls.Add(Me.Trace2Select)
        Me.Controls.Add(Me.Trace2Sel)
        Me.Controls.Add(Me.Trace1Select)
        Me.Controls.Add(Me.Trace1Sel)
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
        Me.Controls.Add(Me.cmbExecute)
        Me.Controls.Add(Me.txtDate2)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimizeBox = False
        Me.Name = "TraceView"
        Me.Text = "Trace View"
        CType(Me.TraceGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmbExecute As System.Windows.Forms.Button
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
    Friend WithEvents Trace1Sel As System.Windows.Forms.ComboBox
    Friend WithEvents Trace1Select As System.Windows.Forms.CheckBox
    Friend WithEvents Trace2Select As System.Windows.Forms.CheckBox
    Friend WithEvents Trace2Sel As System.Windows.Forms.ComboBox
    Friend WithEvents Trace3Select As System.Windows.Forms.CheckBox
    Friend WithEvents Trace3Sel As System.Windows.Forms.ComboBox
    Friend WithEvents Trace4Select As System.Windows.Forms.CheckBox
    Friend WithEvents Trace4Sel As System.Windows.Forms.ComboBox
    Friend WithEvents lblTestID1 As System.Windows.Forms.Label
    Friend WithEvents lblSeialNumber1 As System.Windows.Forms.Label
    Friend WithEvents lblWorkStation1 As System.Windows.Forms.Label
    Friend WithEvents lblPoints1 As System.Windows.Forms.Label
    Friend WithEvents lblDate1 As System.Windows.Forms.Label
    Friend WithEvents txtTestID1 As System.Windows.Forms.Label
    Friend WithEvents txtSerialNumber1 As System.Windows.Forms.Label
    Friend WithEvents txtWorkStation1 As System.Windows.Forms.Label
    Friend WithEvents txtPoints1 As System.Windows.Forms.Label
    Friend WithEvents txtDate1 As System.Windows.Forms.Label
    Friend WithEvents txtWorkStation2 As System.Windows.Forms.Label
    Friend WithEvents lblTestD2 As System.Windows.Forms.Label
    Friend WithEvents lblSeialNumber2 As System.Windows.Forms.Label
    Friend WithEvents lblDate2 As System.Windows.Forms.Label
    Friend WithEvents txtTestID2 As System.Windows.Forms.Label
    Friend WithEvents txtPoints2 As System.Windows.Forms.Label
    Friend WithEvents lblPoints2 As System.Windows.Forms.Label
    Friend WithEvents lblWorkStation2 As System.Windows.Forms.Label
    Friend WithEvents txtSerialNumber2 As System.Windows.Forms.Label
    Friend WithEvents txtSerialNumber3 As System.Windows.Forms.Label
    Friend WithEvents lblPoints3 As System.Windows.Forms.Label
    Friend WithEvents lblWorkStation3 As System.Windows.Forms.Label
    Friend WithEvents txtDate3 As System.Windows.Forms.Label
    Friend WithEvents txtPoints3 As System.Windows.Forms.Label
    Friend WithEvents txtWorkStation3 As System.Windows.Forms.Label
    Friend WithEvents txtTestID3 As System.Windows.Forms.Label
    Friend WithEvents lblDate3 As System.Windows.Forms.Label
    Friend WithEvents lblSeialNumber3 As System.Windows.Forms.Label
    Friend WithEvents lblTestID3 As System.Windows.Forms.Label
    Friend WithEvents txtSerialNumber4 As System.Windows.Forms.Label
    Friend WithEvents lblPoints4 As System.Windows.Forms.Label
    Friend WithEvents lblWorkStation4 As System.Windows.Forms.Label
    Friend WithEvents txtDate4 As System.Windows.Forms.Label
    Friend WithEvents txtPoints4 As System.Windows.Forms.Label
    Friend WithEvents txtWorkStation4 As System.Windows.Forms.Label
    Friend WithEvents txtTestID4 As System.Windows.Forms.Label
    Friend WithEvents lblDate4 As System.Windows.Forms.Label
    Friend WithEvents lblSeialNumber4 As System.Windows.Forms.Label
    Friend WithEvents lblTestID4 As System.Windows.Forms.Label
    Friend WithEvents cmbSerialNumber1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSerialNumber2 As System.Windows.Forms.ComboBox
    Friend WithEvents TraceGrid As System.Windows.Forms.DataGridView
    Friend WithEvents ZG1 As ZedGraph.ZedGraphControl
    Friend WithEvents txtDate2 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
End Class
