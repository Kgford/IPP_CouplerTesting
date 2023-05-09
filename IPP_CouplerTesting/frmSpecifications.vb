
Imports System.Data
Imports System.Data.SqlClient

Public Class frmSpecifications
    Private Index As Integer
    Private ChangeJob As Integer
    Private OldPart As String
    Private OldJOb As String
    Private byp As Boolean = False
  
    Public Sub New()
        Dim SQLstr As String
        Try
            byp = True
            Percent_bypass = True
            InitializeComponent()
            OldPart = Part
            OldJOb = Job

            Me.cmbPart.Items.Clear()
            Me.cmbPart.Items.Add("IPP-")
            SQLstr = "SELECT DISTINCT PartNumber from Specifications "
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    Me.cmbPart.Items.Add(dr.Item(0))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    Me.cmbPart.Items.Add(drLocal.Item(0))
                End While
                atsLocal.Close()
            End If
            Me.cmbPart.SelectedIndex = 0

            Me.cmbSpecType.Items.Clear()
            Me.cmbSpecType.Items.Add("90 DEGREE COUPLER")
            Me.cmbSpecType.Items.Add("90 DEGREE COUPLER SMD")
            Me.cmbSpecType.Items.Add("180 DEGREE COUPLER")
            Me.cmbSpecType.Items.Add("180 DEGREE COUPLER SMD")
            Me.cmbSpecType.Items.Add("BALUN")
            Me.cmbSpecType.Items.Add("BALUN SMD")
            Me.cmbSpecType.Items.Add("BI DIRECTIONAL COUPLER")
            Me.cmbSpecType.Items.Add("BI DIRECTIONAL COUPLER SMD")
            Me.cmbSpecType.Items.Add("SINGLE DIRECTIONAL COUPLER")
            Me.cmbSpecType.Items.Add("SINGLE DIRECTIONAL COUPLER SMD")
            Me.cmbSpecType.Items.Add("DUAL DIRECTIONAL COUPLER")
            Me.cmbSpecType.Items.Add("DUAL DIRECTIONAL COUPLER SMD")
            Me.cmbSpecType.Items.Add("COMBINER/DIVIDER")
            Me.cmbSpecType.Items.Add("COMBINER/DIVIDER SMD")
            Me.cmbSpecType.Items.Add("TRANSFORMER")
            cmbPart.Text = Trim(OldPart)
            cmbJob.Text = Trim(OldJOb)
            Part = OldPart
            Job = OldJOb
            LoadPart()
            PartSpec = Trim(cmbPart.Text)
            jobSpec = Trim(cmbJob.Text)
            byp = True
            txtFail.Text = GetFailPercent()
            If GetBypass() = 1 Then
                sup_bypass.CheckState = CheckState.Checked
            Else
                sup_bypass.CheckState = CheckState.Unchecked
            End If
            byp = False
            Percent_bypass = False
        Catch ex As Exception

        End Try
    End Sub
    Private Sub cmbPart_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPart.SelectedIndexChanged
        LoadSpecType()
        LoadPart()
        Me.Refresh()
        Part = Trim(Me.cmbPart.Text)
        Job = Trim(Me.cmbJob.Text)
    End Sub

    Private Sub cmbPart_Change(sender As Object, e As EventArgs) Handles cmbPart.DataSourceChanged
        If Me.cmbPart.Text = "IPP - " Or Me.cmbPart.Text = "Part Number" Then
            LoadSpecType()
            LoadPart()
            Me.Refresh()
            Part = Trim(Me.cmbPart.Text)
            Job = Trim(Me.cmbJob.Text)
        End If

    End Sub

    Private Sub cmbPart_Click(sender As Object, e As EventArgs) Handles cmbPart.Click
        LoadSpecType()
        LoadPart()
        Me.Refresh()
        Part = Me.cmbPart.Text
        Job = Me.cmbJob.Text
    End Sub

    Private Sub LoadPart()
        Dim SQLstr As String
        SpecAB_TF = False
        ISO_TF = False
        Try
            SQLstr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "'"
            '*******************************************************************************
            'Load the Conguration Data
            '*******************************************************************************
            System.Threading.Thread.Sleep(500)
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(200)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If Index = 0 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(dr.GetValue(1)) Then Me.cmbSpecType.Text = dr.Item(1)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtTitle.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(6)) Then Me.txtStartFreq_1.Text = dr.Item(6)
                        If Not IsDBNull(dr.GetValue(7)) Then Me.txtStopFreq_1.Text = dr.Item(7)
                        If Not IsDBNull(dr.GetValue(20)) Then Me.txtPower_1.Text = dr.Item(20)
                        If Not IsDBNull(dr.GetValue(5)) Then Me.txtQuantity.Text = dr.Item(5)
                        If Not IsDBNull(dr.GetValue(32)) Then Me.txtPPH.Text = dr.Item(32)
                        If Not IsDBNull(dr.GetValue(11)) Then Me.txtTest1_1.Text = dr.Item(11)
                        If Not IsDBNull(dr.GetValue(10)) Then Me.txtTest2_1.Text = dr.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(dr.GetValue(12)) Then Me.txtTest3_1.Text = dr.Item(12)
                            If Not IsDBNull(dr.GetValue(14)) Then Me.txtTest4_1.Text = dr.Item(14)
                            If Not IsDBNull(dr.GetValue(18)) Then Me.txtTest5_1.Text = dr.Item(18)
                            If Not IsDBNull(dr.GetValue(47)) Then
                                If dr.Item(47) = 1 Then
                                    SpecAB_TF = True
                                    If Not IsDBNull(dr.GetValue(42)) Then Me.txtTest4_1_exp.Text = dr.Item(42)
                                    If Not IsDBNull(dr.GetValue(43)) Then Me.tx11Start.Text = dr.Item(43)
                                    If Not IsDBNull(dr.GetValue(44)) Then Me.tx12Start.Text = dr.Item(44)
                                    If Not IsDBNull(dr.GetValue(45)) Then Me.tx11Stop.Text = dr.Item(45)
                                    If Not IsDBNull(dr.GetValue(46)) Then Me.tx12Stop.Text = dr.Item(46)
                                    ck1Expanded.CheckState = CheckState.Checked
                                Else
                                    SpecAB_TF = False
                                    ck1Expanded.CheckState = CheckState.Unchecked
                                End If
                            End If
                        ElseIf Index = 2 Then
                            If Not IsDBNull(dr.GetValue(15)) Then Me.txtTest3_3.Text = dr.Item(15)
                            If Not IsDBNull(dr.GetValue(17)) Then Me.txtTest4_3.Text = dr.Item(17)
                            If Not IsDBNull(dr.GetValue(19)) Then Me.txtTest5_3.Text = dr.Item(19)
                        End If
                        If Not IsDBNull(dr.GetValue(100)) Then SwitchPorts = dr.Item(100)
                        If SwitchPorts = 1 Then
                            ckPorts.CheckState = CheckState.Checked
                        Else
                            ckPorts.CheckState = CheckState.Unchecked
                        End If
                    ElseIf Index = 1 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(dr.GetValue(1)) Then Me.cmbSpecType.Text = dr.Item(1)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtTitle.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(6)) Then Me.txtStartFreq_2.Text = dr.Item(6)
                        If Not IsDBNull(dr.GetValue(7)) Then Me.txtStopFreq_2.Text = dr.Item(7)
                        If Not IsDBNull(dr.GetValue(20)) Then Me.txtPower_2.Text = dr.Item(20)
                        If Not IsDBNull(dr.GetValue(5)) Then Me.txtQuantity.Text = dr.Item(5)
                        If Not IsDBNull(dr.GetValue(32)) Then Me.txtPPH.Text = dr.Item(32)
                        If Not IsDBNull(dr.GetValue(11)) Then Me.txtTest1_2.Text = dr.Item(11)
                        If Not IsDBNull(dr.GetValue(10)) Then Me.txtTest2_2.Text = dr.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(dr.GetValue(14)) Then Me.txtTest4_2.Text = dr.Item(14)
                            If Not IsDBNull(dr.GetValue(18)) Then Me.txtTest5_2.Text = dr.Item(18)
                        ElseIf Index = 2 Then
                            If Not IsDBNull(dr.GetValue(17)) Then Me.txtTest4_2.Text = dr.Item(17)
                            If Not IsDBNull(dr.GetValue(19)) Then Me.txtTest5_2.Text = dr.Item(19)
                        End If
                        If Not IsDBNull(dr.GetValue(100)) Then SwitchPorts = dr.Item(100)
                        If SwitchPorts = 1 Then
                            ckPorts.CheckState = CheckState.Checked
                        Else
                            ckPorts.CheckState = CheckState.Unchecked
                        End If
                    ElseIf Index = 2 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(dr.GetValue(1)) Then Me.cmbSpecType.Text = dr.Item(1)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtTitle.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(6)) Then Me.txtStartFreq_3.Text = dr.Item(6)
                        If Not IsDBNull(dr.GetValue(7)) Then Me.txtStopFreq_3.Text = dr.Item(7)
                        If Not IsDBNull(dr.GetValue(20)) Then Me.txtPower_3.Text = dr.Item(20)
                        If Not IsDBNull(dr.GetValue(5)) Then Me.txtQuantity.Text = dr.Item(5)
                        If Not IsDBNull(dr.GetValue(32)) Then Me.txtPPH.Text = dr.Item(32)
                        If Not IsDBNull(dr.GetValue(16)) Then Me.COUPPlusMinus.Text = dr.Item(16)
                        If Not IsDBNull(dr.GetValue(101)) Then Me.COUPPlus.Text = dr.Item(101)
                        If Not IsDBNull(dr.GetValue(102)) Then Me.COUPMinus.Text = dr.Item(102)
                        If Not IsDBNull(dr.GetValue(103)) Then COupDualSpec = dr.Item(103)
                        If COupDualSpec = True Then
                            ckPlusMinus.CheckState = CheckState.Checked
                        Else
                            ckPlusMinus.CheckState = CheckState.Unchecked
                        End If
                        If Not IsDBNull(dr.GetValue(11)) Then Me.txtTest1_3.Text = dr.Item(11)
                        If Not IsDBNull(dr.GetValue(10)) Then Me.txtTest2_3.Text = dr.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(dr.GetValue(13)) Then Me.txtTest3_1.Text = dr.Item(13)
                            If Not IsDBNull(dr.GetValue(14)) Then Me.txtTest4_1.Text = dr.Item(14)
                            If Not IsDBNull(dr.GetValue(19)) Then Me.txtTest5_1.Text = dr.Item(19)
                            If Not IsDBNull(dr.GetValue(14)) Then Me.txtTest4_2.Text = dr.Item(14)
                            If Not IsDBNull(dr.GetValue(19)) Then Me.txtTest5_2.Text = dr.Item(19)
                            If Not IsDBNull(dr.GetValue(13)) Then Me.txtTest3_4.Text = dr.Item(13)
                            If Not IsDBNull(dr.GetValue(14)) Then Me.txtTest4_4.Text = dr.Item(14)
                            If Not IsDBNull(dr.GetValue(19)) Then Me.txtTest5_4.Text = dr.Item(19)
                        ElseIf Index = 2 Then
                            If Not IsDBNull(dr.GetValue(15)) Then Me.txtTest3_3.Text = dr.Item(15)
                            If Not IsDBNull(dr.GetValue(17)) Then Me.txtTest4_3.Text = dr.Item(17)
                            If Not IsDBNull(dr.GetValue(19)) Then Me.txtTest5_3.Text = dr.Item(19)
                        End If
                        If Not IsDBNull(dr.GetValue(100)) Then SwitchPorts = dr.Item(100)
                        If SwitchPorts = 1 Then
                            ckPorts.CheckState = CheckState.Checked
                        Else
                            ckPorts.CheckState = CheckState.Unchecked
                        End If
                    ElseIf Index = 3 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(dr.GetValue(1)) Then Me.cmbSpecType.Text = dr.Item(1)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtTitle.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(6)) Then Me.txtStartFreq_4.Text = dr.Item(6)
                        If Not IsDBNull(dr.GetValue(7)) Then Me.txtStopFreq_4.Text = dr.Item(7)
                        If Not IsDBNull(dr.GetValue(20)) Then Me.txtPower_4.Text = dr.Item(20)
                        If Not IsDBNull(dr.GetValue(5)) Then Me.txtQuantity.Text = dr.Item(5)
                        If Not IsDBNull(dr.GetValue(32)) Then Me.txtPPH.Text = dr.Item(32)
                        If Not IsDBNull(dr.GetValue(11)) Then Me.txtTest1_4.Text = dr.Item(11)
                        If Not IsDBNull(dr.GetValue(10)) Then Me.txtTest2_4.Text = dr.Item(10)
                        If Not IsDBNull(dr.GetValue(14)) Then Me.txtTest4_4.Text = dr.Item(14)
                        If Not IsDBNull(dr.GetValue(18)) Then Me.txtTest5_4.Text = dr.Item(18)
                        If Not IsDBNull(dr.GetValue(12)) Then Me.txtTest3_4.Text = dr.Item(12)
                        If Not IsDBNull(dr.GetValue(8)) And Not dr.GetValue(8) = 0 Then
                            Me.txtTest3_5.Text = dr.Item(13)
                            ISO_TF = True
                            If Not IsDBNull(dr.GetValue(8)) Or Not dr.GetValue(0) = 0 Then
                                Me.txtISOFreq_1.Text = dr.Item(6)
                                Me.txtISOFreq_2.Text = dr.Item(8)
                                Me.txtISOFreq_3.Text = dr.Item(8)
                                Me.txtISOFreq_4.Text = dr.Item(7)
                                Me.CheckBox2.CheckState = CheckState.Checked
                            End If
                        End If
                        If Not IsDBNull(dr.GetValue(9)) Then Me.txtOutputPorts.Text = dr.Item(9)
                        If Not IsDBNull(dr.GetValue(100)) Then SwitchPorts = dr.Item(100)
                        If SwitchPorts = 1 Then
                            ckPorts.CheckState = CheckState.Checked
                        Else
                            ckPorts.CheckState = CheckState.Unchecked
                        End If
                    ElseIf Index = 4 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(dr.GetValue(1)) Then Me.cmbSpecType.Text = dr.Item(1)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtTitle.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(6)) Then Me.txtStartFreq_5.Text = dr.Item(6)
                        If Not IsDBNull(dr.GetValue(7)) Then Me.txtStopFreq_5.Text = dr.Item(7)
                        If Not IsDBNull(dr.GetValue(20)) Then Me.txtPower_5.Text = dr.Item(20)
                        If Not IsDBNull(dr.GetValue(5)) Then Me.txtQuantity.Text = dr.Item(5)
                        If Not IsDBNull(dr.GetValue(32)) Then Me.txtPPH.Text = dr.Item(32)
                        If Not IsDBNull(dr.GetValue(11)) Then Me.txtTest1_5.Text = dr.Item(11)
                        If Not IsDBNull(dr.GetValue(10)) Then Me.txtTest2_5.Text = dr.Item(10)
                        If Not IsDBNull(dr.GetValue(9)) Then Me.txtOutputPorts.Text = dr.Item(9)
                        If Not IsDBNull(dr.GetValue(96)) Then
                            If dr.Item(91) Then
                                IL_TF = True
                                If Not IsDBNull(dr.GetValue(92)) Then Me.txtTest1_5_exp.Text = dr.Item(92)
                                If Not IsDBNull(dr.GetValue(93)) Then Me.txtILFreq_1.Text = dr.Item(93)
                                If Not IsDBNull(dr.GetValue(94)) Then Me.txtILFreq_3.Text = dr.Item(94)
                                If Not IsDBNull(dr.GetValue(95)) Then Me.txtILFreq_2.Text = dr.Item(95)
                                If Not IsDBNull(dr.GetValue(96)) Then Me.txtILFreq_4.Text = dr.Item(96)
                                ilchk.CheckState = CheckState.Checked
                            Else
                                IL_TF = False
                                ilchk.CheckState = CheckState.Unchecked
                            End If
                        End If
                    End If
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If Index = 0 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(drLocal.GetValue(1)) Then Me.cmbSpecType.Text = drLocal.Item(1)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtTitle.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtStartFreq_1.Text = drLocal.Item(6)
                        If Not IsDBNull(drLocal.GetValue(7)) Then Me.txtStopFreq_1.Text = drLocal.Item(7)
                        If Not IsDBNull(drLocal.GetValue(20)) Then Me.txtPower_1.Text = drLocal.Item(20)
                        If Not IsDBNull(drLocal.GetValue(5)) Then Me.txtQuantity.Text = drLocal.Item(5)
                        If Not IsDBNull(drLocal.GetValue(16)) Then Me.COUPPlusMinus.Text = drLocal.Item(16)
                        If Not IsDBNull(drLocal.GetValue(32)) Then Me.txtPPH.Text = drLocal.Item(32)
                        If Not IsDBNull(drLocal.GetValue(11)) Then Me.txtTest1_1.Text = drLocal.Item(11)
                        If Not IsDBNull(drLocal.GetValue(10)) Then Me.txtTest2_1.Text = drLocal.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(drLocal.GetValue(12)) Then Me.txtTest3_1.Text = drLocal.Item(12)
                            If Not IsDBNull(drLocal.GetValue(14)) Then Me.txtTest4_1.Text = drLocal.Item(14)
                            If Not IsDBNull(drLocal.GetValue(18)) Then Me.txtTest5_1.Text = drLocal.Item(18)
                            If Not IsDBNull(drLocal.GetValue(47)) Then
                                If drLocal.Item(47) = 1 Then
                                    SpecAB_TF = True
                                    If Not IsDBNull(drLocal.GetValue(42)) Then Me.txtTest4_1_exp.Text = drLocal.Item(42)
                                    If Not IsDBNull(drLocal.GetValue(43)) Then Me.tx11Start.Text = drLocal.Item(43)
                                    If Not IsDBNull(drLocal.GetValue(44)) Then Me.tx12Start.Text = drLocal.Item(44)
                                    If Not IsDBNull(drLocal.GetValue(45)) Then Me.tx11Stop.Text = drLocal.Item(45)
                                    If Not IsDBNull(drLocal.GetValue(46)) Then Me.tx12Stop.Text = drLocal.Item(46)
                                    ck1Expanded.CheckState = CheckState.Checked
                                Else
                                    SpecAB_TF = False
                                    ck1Expanded.CheckState = CheckState.Unchecked
                                End If
                            End If
                        ElseIf Index = 2 Then
                            If Not IsDBNull(drLocal.GetValue(15)) Then Me.txtTest3_1.Text = drLocal.Item(15)
                            If Not IsDBNull(drLocal.GetValue(17)) Then Me.txtTest4_1.Text = drLocal.Item(17)
                            If Not IsDBNull(drLocal.GetValue(19)) Then Me.txtTest5_1.Text = drLocal.Item(19)
                        End If
                    ElseIf Index = 1 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(drLocal.GetValue(1)) Then Me.cmbSpecType.Text = drLocal.Item(1)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtTitle.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtStartFreq_2.Text = drLocal.Item(6)
                        If Not IsDBNull(drLocal.GetValue(7)) Then Me.txtStopFreq_2.Text = drLocal.Item(7)
                        If Not IsDBNull(drLocal.GetValue(20)) Then Me.txtPower_2.Text = drLocal.Item(20)
                        If Not IsDBNull(drLocal.GetValue(5)) Then Me.txtQuantity.Text = drLocal.Item(5)
                        If Not IsDBNull(drLocal.GetValue(32)) Then Me.txtPPH.Text = drLocal.Item(32)
                        If Not IsDBNull(drLocal.GetValue(11)) Then Me.txtTest1_2.Text = drLocal.Item(11)
                        If Not IsDBNull(drLocal.GetValue(10)) Then Me.txtTest2_2.Text = drLocal.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(drLocal.GetValue(14)) Then Me.txtTest4_2.Text = drLocal.Item(14)
                            If Not IsDBNull(drLocal.GetValue(18)) Then Me.txtTest5_2.Text = drLocal.Item(18)
                        ElseIf Index = 2 Then
                            If Not IsDBNull(drLocal.GetValue(17)) Then Me.txtTest4_2.Text = drLocal.Item(17)
                            If Not IsDBNull(drLocal.GetValue(19)) Then Me.txtTest5_2.Text = drLocal.Item(19)
                        End If
                    ElseIf Index = 2 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(drLocal.GetValue(1)) Then Me.cmbSpecType.Text = drLocal.Item(1)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtTitle.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtStartFreq_3.Text = drLocal.Item(6)
                        If Not IsDBNull(drLocal.GetValue(7)) Then Me.txtStopFreq_3.Text = drLocal.Item(7)
                        If Not IsDBNull(drLocal.GetValue(20)) Then Me.txtPower_3.Text = drLocal.Item(20)
                        If Not IsDBNull(drLocal.GetValue(5)) Then Me.txtQuantity.Text = drLocal.Item(5)
                        If Not IsDBNull(drLocal.GetValue(32)) Then Me.txtPPH.Text = drLocal.Item(32)
                        If Not IsDBNull(drLocal.GetValue(11)) Then Me.txtTest1_3.Text = drLocal.Item(11)
                        If Not IsDBNull(drLocal.GetValue(10)) Then Me.txtTest2_3.Text = drLocal.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(drLocal.GetValue(12)) Then Me.txtTest3_3.Text = drLocal.Item(12)
                            If Not IsDBNull(drLocal.GetValue(14)) Then Me.txtTest4_3.Text = drLocal.Item(14)
                            If Not IsDBNull(drLocal.GetValue(18)) Then Me.txtTest5_3.Text = drLocal.Item(18)
                        ElseIf Index = 2 Then
                            If Not IsDBNull(drLocal.GetValue(15)) Then Me.txtTest3_3.Text = drLocal.Item(15)
                            If Not IsDBNull(drLocal.GetValue(17)) Then Me.txtTest4_3.Text = drLocal.Item(17)
                            If Not IsDBNull(drLocal.GetValue(19)) Then Me.txtTest5_3.Text = drLocal.Item(19)
                        End If
                    ElseIf Index = 3 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(drLocal.GetValue(1)) Then Me.cmbSpecType.Text = drLocal.Item(1)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtTitle.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtStartFreq_4.Text = drLocal.Item(6)
                        If Not IsDBNull(drLocal.GetValue(7)) Then Me.txtStopFreq_4.Text = drLocal.Item(7)
                        If Not IsDBNull(drLocal.GetValue(20)) Then Me.txtPower_4.Text = drLocal.Item(20)
                        If Not IsDBNull(drLocal.GetValue(5)) Then Me.txtQuantity.Text = drLocal.Item(5)
                        If Not IsDBNull(drLocal.GetValue(32)) Then Me.txtPPH.Text = drLocal.Item(32)
                        If Not IsDBNull(drLocal.GetValue(11)) Then Me.txtTest1_4.Text = drLocal.Item(11)
                        If Not IsDBNull(drLocal.GetValue(10)) Then Me.txtTest2_4.Text = drLocal.Item(10)
                        If Not IsDBNull(drLocal.GetValue(14)) Then Me.txtTest4_4.Text = drLocal.Item(14)
                        If Not IsDBNull(drLocal.GetValue(18)) Then Me.txtTest5_4.Text = drLocal.Item(18)
                        If Not IsDBNull(drLocal.GetValue(12)) Then Me.txtTest3_4.Text = drLocal.Item(12)
                        If Not IsDBNull(drLocal.GetValue(8)) And Not drLocal.GetValue(8) = 0 Then
                            Me.txtTest3_5.Text = drLocal.Item(13)
                            ISO_TF = True
                            If Not IsDBNull(drLocal.GetValue(8)) Or Not drLocal.GetValue(0) = 0 Then
                                Me.txtISOFreq_1.Text = drLocal.Item(6)
                                Me.txtISOFreq_2.Text = drLocal.Item(8)
                                Me.txtISOFreq_3.Text = drLocal.Item(8)
                                Me.txtISOFreq_4.Text = drLocal.Item(7)
                                Me.CheckBox2.CheckState = CheckState.Checked
                            End If
                        End If
                        If Not IsDBNull(drLocal.GetValue(9)) Then Me.txtOutputPorts.Text = drLocal.Item(9)
                    ElseIf Index = 5 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(drLocal.GetValue(1)) Then Me.cmbSpecType.Text = drLocal.Item(1)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtTitle.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtStartFreq_5.Text = drLocal.Item(6)
                        If Not IsDBNull(drLocal.GetValue(7)) Then Me.txtStopFreq_5.Text = drLocal.Item(7)
                        If Not IsDBNull(drLocal.GetValue(20)) Then Me.txtPower_5.Text = drLocal.Item(20)
                        If Not IsDBNull(drLocal.GetValue(5)) Then Me.txtQuantity.Text = drLocal.Item(5)
                        If Not IsDBNull(drLocal.GetValue(32)) Then Me.txtPPH.Text = drLocal.Item(32)
                        If Not IsDBNull(drLocal.GetValue(11)) Then Me.txtTest1_5.Text = drLocal.Item(11)
                        If Not IsDBNull(drLocal.GetValue(10)) Then Me.txtTest2_5.Text = drLocal.Item(10)
                        If Not IsDBNull(drLocal.GetValue(9)) Then Me.txtOutputPorts.Text = drLocal.Item(9)
                        If Not IsDBNull(drLocal.GetValue(47)) Then
                            If drLocal.Item(97) = 1 Then
                                IL_TF = True
                                If Not IsDBNull(drLocal.GetValue(92)) Then Me.txtTest4_1_exp.Text = drLocal.Item(92)
                                If Not IsDBNull(drLocal.GetValue(93)) Then Me.tx11Start.Text = drLocal.Item(93)
                                If Not IsDBNull(drLocal.GetValue(94)) Then Me.tx12Start.Text = drLocal.Item(94)
                                If Not IsDBNull(drLocal.GetValue(95)) Then Me.tx11Stop.Text = drLocal.Item(95)
                                If Not IsDBNull(drLocal.GetValue(96)) Then Me.tx12Stop.Text = drLocal.Item(96)
                                ilchk.CheckState = CheckState.Checked
                            Else
                                IL_TF = False
                                ilchk.CheckState = CheckState.Unchecked
                            End If
                        End If
                    End If
                End While
            End If

            If Index = 0 And cmbPart.Text.Contains("IT") Then
                Me.txtJ1J1_1.Text = "----"
                Me.txtJ1J2_1.Text = "ISOLATION"
                Me.txtJ1J3_1.Text = "-6dB < 0 DEG"
                Me.txtJ1J4_1.Text = "-6dB < -90 DEG"
                Me.txtJ2J1_1.Text = "ISOLATION"
                Me.txtJ2J2_1.Text = "----"
                Me.txtJ2J3_1.Text = "-6dB < -90 DEG"
                Me.txtJ2J4_1.Text = "-6dB < 0 DEG"
                Me.txtJ3J1_1.Text = "-6dB < 0 DEG"
                Me.txtJ3J2_1.Text = "-6dB < -90 DEG"
                Me.txtJ3J3_1.Text = "----"
                Me.txtJ3J4_1.Text = "ISOLATION"
                Me.txtJ4J1_1.Text = "-6dB < -90 DEG"
                Me.txtJ4J2_1.Text = "-6dB < 0 DEG"
                Me.txtJ4J3_1.Text = "ISOLATION"
                Me.txtJ4J4_1.Text = "----"
            ElseIf Index = 0 And Me.cmbSpecType.Text = "180 DEGREE COUPLER" Or Me.cmbSpecType.Text = "180 DEGREE COUPLER SMD" Then
                Label5.Text = "180 DEGREE COUPLER"
                Me.txtJ1J1_1.Text = "----"
                Me.txtJ1J2_1.Text = "ISOLATION"
                Me.txtJ1J3_1.Text = "-3dB < 0 DEG"
                Me.txtJ1J4_1.Text = "-3dB < -90 DEG"
                Me.txtJ2J1_1.Text = "ISOLATION"
                Me.txtJ2J2_1.Text = "----"
                Me.txtJ2J3_1.Text = "-3dB < -180 DEG"
                Me.txtJ2J4_1.Text = "-3dB < 0 DEG"
                Me.txtJ3J1_1.Text = "-3dB < 0 DEG"
                Me.txtJ3J2_1.Text = "-3dB < -90 DEG"
                Me.txtJ3J3_1.Text = "----"
                Me.txtJ3J4_1.Text = "ISOLATION"
                Me.txtJ4J1_1.Text = "-3dB < -90 DEG"
                Me.txtJ4J2_1.Text = "-3dB< 0 DEG"
                Me.txtJ4J3_1.Text = "ISOLATION"
                Me.txtJ4J4_1.Text = "----"
            ElseIf Index = 0 Then
                Me.txtJ1J1_1.Text = "----"
                Me.txtJ1J2_1.Text = "ISOLATION"
                Me.txtJ1J3_1.Text = "-3dB < 0 DEG"
                Me.txtJ1J4_1.Text = "-3dB < -90 DEG"
                Me.txtJ2J1_1.Text = "ISOLATION"
                Me.txtJ2J2_1.Text = "----"
                Me.txtJ2J3_1.Text = "-3dB < -90 DEG"
                Me.txtJ2J4_1.Text = "-3dB < 0 DEG"
                Me.txtJ3J1_1.Text = "-3dB < 0 DEG"
                Me.txtJ3J2_1.Text = "-3dB < -90 DEG"
                Me.txtJ3J3_1.Text = "----"
                Me.txtJ3J4_1.Text = "ISOLATION"
                Me.txtJ4J1_1.Text = "-3dB < -90 DEG"
                Me.txtJ4J2_1.Text = "-3dB< 0 DEG"
                Me.txtJ4J3_1.Text = "ISOLATION"
                Me.txtJ4J4_1.Text = "----"
            ElseIf Index = 1 Then
                Me.txtJ1J1_2.Text = "----"
                Me.txtJ1J2_2.Text = "-6dB < 0 DEG"
                Me.txtJ1J3_2.Text = "-6dB < -180 DEG"
            End If

            SQLstr = "SELECT * from PortConfig where PartNumber = '" & Me.cmbPart.Text & "'"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If Index = 0 Then
                        If Not IsDBNull(dr.GetValue(3)) Then Me.txtJ1J1_1.Text = dr.Item(3)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtJ1J2_1.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(5)) Then Me.txtJ1J3_1.Text = dr.Item(5)
                        If Not IsDBNull(dr.GetValue(6)) Then Me.txtJ1J4_1.Text = dr.Item(6)
                        If Not IsDBNull(dr.GetValue(8)) Then Me.txtJ2J1_1.Text = dr.Item(8)
                        If Not IsDBNull(dr.GetValue(9)) Then Me.txtJ2J2_1.Text = dr.Item(9)
                        If Not IsDBNull(dr.GetValue(10)) Then Me.txtJ2J3_1.Text = dr.Item(10)
                        If Not IsDBNull(dr.GetValue(11)) Then Me.txtJ2J4_1.Text = dr.Item(11)
                        If Not IsDBNull(dr.GetValue(12)) Then Me.txtJ3J1_1.Text = dr.Item(12)
                        If Not IsDBNull(dr.GetValue(13)) Then Me.txtJ3J2_1.Text = dr.Item(13)
                        If Not IsDBNull(dr.GetValue(14)) Then Me.txtJ3J3_1.Text = dr.Item(14)
                        If Not IsDBNull(dr.GetValue(15)) Then Me.txtJ3J4_1.Text = dr.Item(15)
                        If Not IsDBNull(dr.GetValue(16)) Then Me.txtJ4J1_1.Text = dr.Item(16)
                        If Not IsDBNull(dr.GetValue(17)) Then Me.txtJ4J2_1.Text = dr.Item(17)
                        If Not IsDBNull(dr.GetValue(18)) Then Me.txtJ4J3_1.Text = dr.Item(18)
                        If Not IsDBNull(dr.GetValue(19)) Then Me.txtJ4J4_1.Text = dr.Item(19)
                    ElseIf Index = 1 Then
                        If Not IsDBNull(dr.GetValue(3)) Then Me.txtJ1J1_2.Text = dr.Item(3)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtJ1J2_2.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(5)) Then Me.txtJ1J3_2.Text = dr.Item(5)
                    End If

                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If Index = 0 Then
                        If Not IsDBNull(drLocal.GetValue(3)) Then Me.txtJ1J1_1.Text = drLocal.Item(3)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtJ1J2_1.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(5)) Then Me.txtJ1J3_1.Text = drLocal.Item(5)
                        If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtJ1J4_1.Text = drLocal.Item(6)
                        If Not IsDBNull(drLocal.GetValue(8)) Then Me.txtJ2J1_1.Text = drLocal.Item(8)
                        If Not IsDBNull(drLocal.GetValue(9)) Then Me.txtJ2J2_1.Text = drLocal.Item(9)
                        If Not IsDBNull(drLocal.GetValue(10)) Then Me.txtJ2J3_1.Text = drLocal.Item(10)
                        If Not IsDBNull(drLocal.GetValue(11)) Then Me.txtJ2J4_1.Text = drLocal.Item(11)
                        If Not IsDBNull(drLocal.GetValue(12)) Then Me.txtJ3J1_1.Text = drLocal.Item(12)
                        If Not IsDBNull(drLocal.GetValue(13)) Then Me.txtJ3J2_1.Text = drLocal.Item(13)
                        If Not IsDBNull(drLocal.GetValue(14)) Then Me.txtJ3J3_1.Text = drLocal.Item(14)
                        If Not IsDBNull(drLocal.GetValue(15)) Then Me.txtJ3J4_1.Text = drLocal.Item(15)
                        If Not IsDBNull(drLocal.GetValue(16)) Then Me.txtJ4J1_1.Text = drLocal.Item(16)
                        If Not IsDBNull(drLocal.GetValue(17)) Then Me.txtJ4J2_1.Text = drLocal.Item(17)
                        If Not IsDBNull(drLocal.GetValue(18)) Then Me.txtJ4J3_1.Text = drLocal.Item(18)
                        If Not IsDBNull(drLocal.GetValue(19)) Then Me.txtJ4J4_1.Text = drLocal.Item(19)
                    ElseIf Index = 1 Then
                        If Not IsDBNull(drLocal.GetValue(3)) Then Me.txtJ1J1_2.Text = drLocal.Item(3)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtJ1J2_2.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(5)) Then Me.txtJ1J3_2.Text = drLocal.Item(5)
                    End If
                End While
                atsLocal.Close()
            End If
            ChangeJob = 0
            Exit Sub
        Catch ex As Exception
            ChangeJob = 0
        End Try
    End Sub
    Private Sub cmbJob_Change(sender As Object, e As EventArgs) Handles cmbSpecType.DataSourceChanged
        If Not Me.cmbPart.Text = "IPP - " Or Not Me.cmbPart.Text = "Part Number" Then
            If Not Me.cmbJob.Text = "" Then LoadJob()
            Me.Refresh()
            If Not IsNothing(Part) Then Part = Me.cmbPart.Text
            If Not IsNothing(Job) Then Job = Me.cmbJob.Text
        End If
    End Sub
    Private Sub cmbJob_Click(sender As Object, e As EventArgs) Handles cmbJob.Click
        If Not Me.cmbJob.Text = "" Then LoadJob()
        'Me.Refresh()
        If Not IsNothing(Part) Then Part = Me.cmbPart.Text
        If Not IsNothing(Job) Then Job = Me.cmbJob.Text
    End Sub

    Private Sub LoadJob()
        Dim SQLstr As String
        Try
            SQLstr = "SELECT * from Specifications where JobNumber = '" & Me.cmbJob.Text & "'"

            '*******************************************************************************
            'Load the Conguration Data
            '*******************************************************************************
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If Index = 0 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(dr.GetValue(1)) Then Me.cmbSpecType.Text = dr.Item(1)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtTitle.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(6)) Then Me.txtStartFreq_1.Text = dr.Item(6)
                        If Not IsDBNull(dr.GetValue(7)) Then Me.txtStopFreq_1.Text = dr.Item(7)
                        If Not IsDBNull(dr.GetValue(20)) Then Me.txtPower_1.Text = dr.Item(20)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtQuantity.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(32)) Then Me.txtPPH.Text = dr.Item(32)
                        If Not IsDBNull(dr.GetValue(11)) Then Me.txtTest1_1.Text = dr.Item(11)
                        If Not IsDBNull(dr.GetValue(10)) Then Me.txtTest2_1.Text = dr.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(dr.GetValue(12)) Then Me.txtTest3_1.Text = dr.Item(12)
                            If Not IsDBNull(dr.GetValue(14)) Then Me.txtTest4_1.Text = dr.Item(14)
                            If Not IsDBNull(dr.GetValue(18)) Then Me.txtTest5_1.Text = dr.Item(18)
                            If Not IsDBNull(dr.GetValue(47)) Then
                                If dr.Item(47) = 1 Then
                                    SpecAB_TF = True
                                    If Not IsDBNull(dr.GetValue(42)) Then Me.txtTest4_1_exp.Text = dr.Item(42)
                                    If Not IsDBNull(dr.GetValue(43)) Then Me.tx11Start.Text = dr.Item(43)
                                    If Not IsDBNull(dr.GetValue(44)) Then Me.tx12Start.Text = dr.Item(44)
                                    If Not IsDBNull(dr.GetValue(45)) Then Me.tx11Stop.Text = dr.Item(45)
                                    If Not IsDBNull(dr.GetValue(46)) Then Me.tx12Stop.Text = dr.Item(46)
                                    ck1Expanded.CheckState = CheckState.Checked
                                Else
                                    SpecAB_TF = False
                                    ck1Expanded.CheckState = CheckState.Unchecked
                                End If
                            End If
                        ElseIf Index = 2 Then
                            If Not IsDBNull(dr.GetValue(15)) Then Me.txtTest3_3.Text = dr.Item(15)
                            If Not IsDBNull(dr.GetValue(17)) Then Me.txtTest4_3.Text = dr.Item(17)
                            If Not IsDBNull(dr.GetValue(19)) Then Me.txtTest5_3.Text = dr.Item(19)
                        End If
                        If Not IsDBNull(dr.GetValue(100)) Then SwitchPorts = dr.Item(100)
                        If SwitchPorts = 1 Then
                            ckPorts.CheckState = CheckState.Checked
                        Else
                            ckPorts.CheckState = CheckState.Unchecked
                        End If
                    ElseIf Index = 1 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(dr.GetValue(1)) Then Me.cmbSpecType.Text = dr.Item(1)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtTitle.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(6)) Then Me.txtStartFreq_2.Text = dr.Item(6)
                        If Not IsDBNull(dr.GetValue(7)) Then Me.txtStopFreq_2.Text = dr.Item(7)
                        If Not IsDBNull(dr.GetValue(20)) Then Me.txtPower_2.Text = dr.Item(20)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtQuantity.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(32)) Then Me.txtPPH.Text = dr.Item(32)
                        If Not IsDBNull(dr.GetValue(11)) Then Me.txtTest1_2.Text = dr.Item(11)
                        If Not IsDBNull(dr.GetValue(10)) Then Me.txtTest2_2.Text = dr.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(dr.GetValue(14)) Then Me.txtTest4_2.Text = dr.Item(14)
                            If Not IsDBNull(dr.GetValue(18)) Then Me.txtTest5_2.Text = dr.Item(18)
                        ElseIf Index = 2 Then
                            If Not IsDBNull(dr.GetValue(17)) Then Me.txtTest4_2.Text = dr.Item(17)
                            If Not IsDBNull(dr.GetValue(19)) Then Me.txtTest5_2.Text = dr.Item(19)
                        End If
                        If Not IsDBNull(dr.GetValue(100)) Then SwitchPorts = dr.Item(100)
                        If SwitchPorts = 1 Then
                            ckPorts.CheckState = CheckState.Checked
                        Else
                            ckPorts.CheckState = CheckState.Unchecked
                        End If
                    ElseIf Index = 2 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(dr.GetValue(1)) Then Me.cmbSpecType.Text = dr.Item(1)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtTitle.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(6)) Then Me.txtStartFreq_3.Text = dr.Item(6)
                        If Not IsDBNull(dr.GetValue(20)) Then Me.txtPower_3.Text = dr.Item(20)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtQuantity.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(32)) Then Me.txtPPH.Text = dr.Item(32)
                        If Not IsDBNull(dr.GetValue(11)) Then Me.txtTest1_3.Text = dr.Item(11)
                        If Not IsDBNull(dr.GetValue(10)) Then Me.txtTest2_3.Text = dr.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(dr.GetValue(13)) Then Me.txtTest3_1.Text = dr.Item(13)
                            If Not IsDBNull(dr.GetValue(14)) Then Me.txtTest4_1.Text = dr.Item(14)
                            If Not IsDBNull(dr.GetValue(19)) Then Me.txtTest5_1.Text = dr.Item(19)
                            'If Not IsDBNull(dr.GetValue(13)) Then Me.txtTest3_2.Text = dr.Item(13)
                            If Not IsDBNull(dr.GetValue(14)) Then Me.txtTest4_2.Text = dr.Item(14)
                            If Not IsDBNull(dr.GetValue(19)) Then Me.txtTest5_2.Text = dr.Item(19)
                            If Not IsDBNull(dr.GetValue(13)) Then Me.txtTest3_4.Text = dr.Item(13)
                            If Not IsDBNull(dr.GetValue(14)) Then Me.txtTest4_4.Text = dr.Item(14)
                            If Not IsDBNull(dr.GetValue(19)) Then Me.txtTest5_4.Text = dr.Item(19)
                        ElseIf Index = 2 Then
                            If Not IsDBNull(dr.GetValue(16)) Then Me.txtTest3_3.Text = dr.Item(15)
                            If Not IsDBNull(dr.GetValue(17)) Then Me.txtTest4_3.Text = dr.Item(17)
                            If Not IsDBNull(dr.GetValue(20)) Then Me.txtTest5_3.Text = dr.Item(20)
                        End If
                        If Not IsDBNull(dr.GetValue(100)) Then SwitchPorts = dr.Item(100)
                        If SwitchPorts = 1 Then
                            ckPorts.CheckState = CheckState.Checked
                        Else
                            ckPorts.CheckState = CheckState.Unchecked
                        End If
                    ElseIf Index = 5 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(dr.GetValue(1)) Then Me.cmbSpecType.Text = dr.Item(1)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtTitle.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(6)) Then Me.txtStartFreq_3.Text = dr.Item(6)
                        If Not IsDBNull(dr.GetValue(20)) Then Me.txtPower_3.Text = dr.Item(20)
                        If Not IsDBNull(dr.GetValue(4)) Then Me.txtQuantity.Text = dr.Item(4)
                        If Not IsDBNull(dr.GetValue(32)) Then Me.txtPPH.Text = dr.Item(32)
                        If Not IsDBNull(dr.GetValue(11)) Then Me.txtTest1_5.Text = dr.Item(11)
                        If Not IsDBNull(dr.GetValue(10)) Then Me.txtTest2_5.Text = dr.Item(10)
                    End If
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If Index = 0 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(drLocal.GetValue(1)) Then Me.cmbSpecType.Text = drLocal.Item(1)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtTitle.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtStartFreq_1.Text = drLocal.Item(6)
                        If Not IsDBNull(drLocal.GetValue(7)) Then Me.txtStopFreq_1.Text = drLocal.Item(7)
                        If Not IsDBNull(drLocal.GetValue(20)) Then Me.txtPower_1.Text = drLocal.Item(20)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtQuantity.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(32)) Then Me.txtPPH.Text = drLocal.Item(32)
                        If Not IsDBNull(drLocal.GetValue(11)) Then Me.txtTest1_1.Text = drLocal.Item(11)
                        If Not IsDBNull(drLocal.GetValue(10)) Then Me.txtTest2_1.Text = drLocal.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(drLocal.GetValue(12)) Then Me.txtTest3_1.Text = drLocal.Item(12)
                            If Not IsDBNull(drLocal.GetValue(14)) Then Me.txtTest4_1.Text = drLocal.Item(14)
                            If Not IsDBNull(drLocal.GetValue(18)) Then Me.txtTest5_1.Text = drLocal.Item(18)
                            If Not IsDBNull(drLocal.GetValue(47)) Then
                                If drLocal.Item(47) = 1 Then
                                    SpecAB_TF = True
                                    If Not IsDBNull(drLocal.GetValue(42)) Then Me.txtTest4_1_exp.Text = drLocal.Item(42)
                                    If Not IsDBNull(drLocal.GetValue(43)) Then Me.tx11Start.Text = drLocal.Item(43)
                                    If Not IsDBNull(drLocal.GetValue(44)) Then Me.tx12Start.Text = drLocal.Item(44)
                                    If Not IsDBNull(drLocal.GetValue(45)) Then Me.tx11Stop.Text = drLocal.Item(45)
                                    If Not IsDBNull(drLocal.GetValue(46)) Then Me.tx12Stop.Text = drLocal.Item(46)
                                    ck1Expanded.CheckState = CheckState.Checked
                                Else
                                    SpecAB_TF = False
                                    ck1Expanded.CheckState = CheckState.Unchecked
                                End If
                            End If
                        ElseIf Index = 2 Then
                            If Not IsDBNull(drLocal.GetValue(15)) Then Me.txtTest3_1.Text = drLocal.Item(15)
                            If Not IsDBNull(drLocal.GetValue(17)) Then Me.txtTest4_1.Text = drLocal.Item(17)
                            If Not IsDBNull(drLocal.GetValue(19)) Then Me.txtTest5_1.Text = drLocal.Item(19)
                        End If
                    ElseIf Index = 1 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(drLocal.GetValue(1)) Then Me.cmbSpecType.Text = drLocal.Item(1)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtTitle.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtStartFreq_2.Text = drLocal.Item(6)
                        If Not IsDBNull(drLocal.GetValue(7)) Then Me.txtStopFreq_2.Text = drLocal.Item(7)
                        If Not IsDBNull(drLocal.GetValue(20)) Then Me.txtPower_2.Text = drLocal.Item(20)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtQuantity.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(32)) Then Me.txtPPH.Text = drLocal.Item(32)
                        If Not IsDBNull(drLocal.GetValue(11)) Then Me.txtTest1_2.Text = drLocal.Item(11)
                        If Not IsDBNull(drLocal.GetValue(10)) Then Me.txtTest2_2.Text = drLocal.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(drLocal.GetValue(14)) Then Me.txtTest4_2.Text = drLocal.Item(14)
                            If Not IsDBNull(drLocal.GetValue(18)) Then Me.txtTest5_2.Text = drLocal.Item(18)
                        ElseIf Index = 2 Then
                            If Not IsDBNull(drLocal.GetValue(17)) Then Me.txtTest4_2.Text = drLocal.Item(17)
                            If Not IsDBNull(drLocal.GetValue(19)) Then Me.txtTest5_2.Text = drLocal.Item(19)
                        End If
                    ElseIf Index = 2 Then
                        If Me.cmbJob.Text = "Job Number" Or ChangeJob = 1 Then Me.cmbJob.Text = ""
                        If Not IsDBNull(drLocal.GetValue(1)) Then Me.cmbSpecType.Text = drLocal.Item(1)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtTitle.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtStartFreq_3.Text = drLocal.Item(6)
                        If Not IsDBNull(drLocal.GetValue(20)) Then Me.txtPower_3.Text = drLocal.Item(20)
                        If Not IsDBNull(drLocal.GetValue(4)) Then Me.txtQuantity.Text = drLocal.Item(4)
                        If Not IsDBNull(drLocal.GetValue(32)) Then Me.txtPPH.Text = drLocal.Item(32)
                        If Not IsDBNull(drLocal.GetValue(11)) Then Me.txtTest1_3.Text = drLocal.Item(11)
                        If Not IsDBNull(drLocal.GetValue(10)) Then Me.txtTest2_3.Text = drLocal.Item(10)
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Not IsDBNull(drLocal.GetValue(12)) Then Me.txtTest3_3.Text = drLocal.Item(12)
                            If Not IsDBNull(drLocal.GetValue(14)) Then Me.txtTest4_3.Text = drLocal.Item(14)
                            If Not IsDBNull(drLocal.GetValue(18)) Then Me.txtTest5_3.Text = drLocal.Item(18)
                        ElseIf Index = 2 Then
                            If Not IsDBNull(drLocal.GetValue(15)) Then Me.txtTest3_3.Text = drLocal.Item(15)
                            If Not IsDBNull(drLocal.GetValue(17)) Then Me.txtTest4_3.Text = drLocal.Item(17)
                            If Not IsDBNull(drLocal.GetValue(19)) Then Me.txtTest5_3.Text = drLocal.Item(19)
                        End If
                    End If
                End While
            End If
            If Index = 0 Then
                Me.txtJ1J1_1.Text = "----"
                Me.txtJ1J2_1.Text = "ISOLATION"
                Me.txtJ1J3_1.Text = "-3dB < 0 DEG"
                Me.txtJ1J4_1.Text = "-3dB < -90 DEG"
                Me.txtJ2J1_1.Text = "ISOLATION"
                Me.txtJ2J2_1.Text = "----"
                Me.txtJ2J3_1.Text = "-3dB < -90 DEG"
                Me.txtJ2J4_1.Text = "-3dB < 0 DEG"
                Me.txtJ3J1_1.Text = "-3dB < 0 DEG"
                Me.txtJ3J2_1.Text = "-3dB < -90 DEG"
                Me.txtJ3J3_1.Text = "----"
                Me.txtJ3J4_1.Text = "ISOLATION"
                Me.txtJ4J1_1.Text = "-3dB < -90 DEG"
                Me.txtJ4J2_1.Text = "-3dB < 0 DEG"
                Me.txtJ4J3_1.Text = "ISOLATION"
                Me.txtJ4J4_1.Text = "----"
            ElseIf Index = 1 Then
                Me.txtJ1J1_2.Text = "----"
                Me.txtJ1J2_2.Text = "-6dB < 0 DEG"
                Me.txtJ1J3_2.Text = "-6dB < -180 DEG"
            End If
            SQLstr = "SELECT * from PortConfig where PartNumber = '" & Me.cmbPart.Text & "'"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If Index = 0 Then
                        Me.txtJ1J1_1.Text = dr.Item(3)
                        Me.txtJ1J2_1.Text = dr.Item(4)
                        Me.txtJ1J3_1.Text = dr.Item(5)
                        Me.txtJ1J4_1.Text = dr.Item(6)
                        Me.txtJ2J1_1.Text = dr.Item(8)
                        Me.txtJ2J2_1.Text = dr.Item(9)
                        Me.txtJ2J3_1.Text = dr.Item(10)
                        Me.txtJ2J4_1.Text = dr.Item(11)
                        Me.txtJ3J1_1.Text = dr.Item(12)
                        Me.txtJ3J2_1.Text = dr.Item(13)
                        Me.txtJ3J3_1.Text = dr.Item(14)
                        Me.txtJ3J4_1.Text = dr.Item(15)
                        Me.txtJ4J1_1.Text = dr.Item(16)
                        Me.txtJ4J2_1.Text = dr.Item(17)
                        Me.txtJ4J3_1.Text = dr.Item(18)
                        Me.txtJ4J4_1.Text = dr.Item(19)
                    ElseIf Index = 1 Then
                        Me.txtJ1J1_2.Text = dr.Item(3)
                        Me.txtJ1J2_2.Text = dr.Item(4)
                        Me.txtJ1J3_2.Text = dr.Item(5)
                    End If
                    ats.Close()
                End While
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If Index = 0 Then
                        Me.txtJ1J1_1.Text = drLocal.Item(3)
                        Me.txtJ1J2_1.Text = drLocal.Item(4)
                        Me.txtJ1J3_1.Text = drLocal.Item(5)
                        Me.txtJ1J4_1.Text = drLocal.Item(6)
                        Me.txtJ2J1_1.Text = drLocal.Item(8)
                        Me.txtJ2J2_1.Text = drLocal.Item(9)
                        Me.txtJ2J3_1.Text = drLocal.Item(10)
                        Me.txtJ2J4_1.Text = drLocal.Item(11)
                        Me.txtJ3J1_1.Text = drLocal.Item(12)
                        Me.txtJ3J2_1.Text = drLocal.Item(13)
                        Me.txtJ3J3_1.Text = drLocal.Item(14)
                        Me.txtJ3J4_1.Text = drLocal.Item(15)
                        Me.txtJ4J1_1.Text = drLocal.Item(16)
                        Me.txtJ4J2_1.Text = drLocal.Item(17)
                        Me.txtJ4J3_1.Text = drLocal.Item(18)
                        Me.txtJ4J4_1.Text = drLocal.Item(19)
                    ElseIf Index = 1 Then
                        Me.txtJ1J1_2.Text = drLocal.Item(3)
                        Me.txtJ1J2_2.Text = drLocal.Item(4)
                        Me.txtJ1J3_2.Text = drLocal.Item(5)
                    End If
                    atsLocal.Close()
                End While
            End If
            ChangeJob = 0
            Exit Sub
        Catch ex As Exception
            ChangeJob = 0
        End Try
    End Sub

    Private Sub cmbSpecType_Change(sender As Object, e As EventArgs) Handles cmbSpecType.DataSourceChanged
        LoadSpecType()
    End Sub

    Private Sub cmbSpecType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSpecType.SelectedIndexChanged
        LoadSpecType()
    End Sub

    Private Sub LoadSpecType()
        If Me.cmbSpecType.Text = "90 DEGREE COUPLER" Or Me.cmbSpecType.Text = "90 DEGREE COUPLER SMD" Or Me.cmbSpecType.Text = "180 DEGREE COUPLER" Or Me.cmbSpecType.Text = "180 DEGREE COUPLER SMD" Then
            Frame1.Visible = True
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = False
            Frame5.Visible = False
            txtTitle.Text = cmbSpecType.Text
            Index = 0
            If Me.cmbSpecType.Text = "180 DEGREE COUPLER" Or Me.cmbSpecType.Text = "180 DEGREE COUPLER SMD" Then
                Label5.Text = "180 DEGREE COUPLER"
                Me.txtJ1J1_1.Text = "----"
                Me.txtJ1J2_1.Text = "ISOLATION"
                Me.txtJ1J3_1.Text = "-3dB < 0 DEG"
                Me.txtJ1J4_1.Text = "-3dB < -90 DEG"
                Me.txtJ2J1_1.Text = "ISOLATION"
                Me.txtJ2J2_1.Text = "----"
                Me.txtJ2J3_1.Text = "-3dB < -180 DEG"
                Me.txtJ2J4_1.Text = "-3dB < 0 DEG"
                Me.txtJ3J1_1.Text = "-3dB < 0 DEG"
                Me.txtJ3J2_1.Text = "-3dB < -90 DEG"
                Me.txtJ3J3_1.Text = "----"
                Me.txtJ3J4_1.Text = "ISOLATION"
                Me.txtJ4J1_1.Text = "-3dB < -90 DEG"
                Me.txtJ4J2_1.Text = "-3dB< 0 DEG"
                Me.txtJ4J3_1.Text = "ISOLATION"
                Me.txtJ4J4_1.Text = "----"
            End If

        ElseIf Me.cmbSpecType.Text = "BALUN" Or Me.cmbSpecType.Text = "BALUN SMD" Then
            Frame1.Visible = False
            Frame2.Visible = True
            Frame3.Visible = False
            Frame4.Visible = False
            Frame5.Visible = False
            txtTitle.Text = cmbSpecType.Text
            Index = 1
        ElseIf Me.cmbSpecType.Text = "BI DIRECTIONAL COUPLER" Or Me.cmbSpecType.Text = "BI DIRECTIONAL COUPLER SMD" Then
            Frame1.Visible = False
            Frame2.Visible = False
            Frame3.Visible = True
            Frame4.Visible = False
            Frame5.Visible = False
            txtTitle.Text = cmbSpecType.Text
            Index = 2
        ElseIf Me.cmbSpecType.Text = "SINGLE DIRECTIONAL COUPLER" Or Me.cmbSpecType.Text = "SINGLE DIRECTIONAL COUPLER SMD" Then
            Frame1.Visible = False
            Frame2.Visible = False
            Frame3.Visible = True
            Frame4.Visible = False
            Frame5.Visible = False
            txtTitle.Text = cmbSpecType.Text
            Index = 2
        ElseIf Me.cmbSpecType.Text = "DUAL DIRECTIONAL COUPLER" Or Me.cmbSpecType.Text = "DUAL DIRECTIONAL COUPLER SMD" Then
            Frame1.Visible = False
            Frame2.Visible = False
            Frame3.Visible = True
            Frame4.Visible = False
            Frame5.Visible = False
            txtTitle.Text = cmbSpecType.Text
            Index = 2
        ElseIf cmbSpecType.Text = "COMBINER/DIVIDER" Or cmbSpecType.Text = "COMBINER/DIVIDER SMD" Then
            Frame1.Visible = False
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = True
            Frame5.Visible = False
            txtTitle.Text = cmbSpecType.Text
            Index = 3
        ElseIf cmbSpecType.Text = "TRANSFORMER" Or cmbSpecType.Text = "TRANSFORMER SMD" Then
            Frame1.Visible = False
            Frame2.Visible = False
            Frame3.Visible = False
            Frame4.Visible = False
            Frame5.Visible = True
            txtTitle.Text = cmbSpecType.Text
            Index = 4
        End If
    End Sub

    Private Sub cmdSaveDatabase_Click(sender As Object, e As EventArgs) Handles cmdSaveDatabase.Click
        Dim SQLstr As String
        Dim Expression As String

        Try

            If Me.cmbJob.Text = "" Then GoTo StupidUser
            Select Case Index
                Case 1
                    If InStr(Me.txtTest2_1.Text, ":") Then GoTo NoRatio
                Case 2
                    If InStr(Me.txtTest2_2.Text, ":") Then GoTo NoRatio
                Case 3
                    If InStr(Me.txtTest2_3.Text, ":") Then GoTo NoRatio
                Case 4
                    If InStr(Me.txtTest2_4.Text, ":") Then GoTo NoRatio
                Case 5
                    If InStr(Me.txtTest2_5.Text, ":") Then GoTo NoRatio
            End Select
            If sup_bypass.CheckState = CheckState.Checked Then
                M_bypass = 1
            Else
                M_bypass = 0
            End If

            Expression = " where PartNumber = '" & Trim(Me.cmbPart.Text) & "' And JobNumber = '" & Trim(Me.cmbJob.Text) & "'"
            SQLstr = "SELECT * from Specifications" & Expression
            If SQL.CheckforRow(SQLstr, "LocalSpecs") = 0 Then
                SQLstr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Trim(Me.cmbJob.Text) & "','" & Trim(Me.cmbPart.Text) & "')"
                SQL.ExecuteSQLCommand(SQLstr, "LocalSpecs")
            End If

            '"~~>************************SQL Or Access Conection****************************
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                Select Case Index
                    Case 0
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Me.txtTest3_1.Text = "" Or Me.txtTest4_1.Text = "" Or Me.txtTest5_1.Text = "" Then GoTo StupidUser
                            If Not IsNumeric(Me.txtTest3_1.Text) Or Not IsNumeric(Me.txtTest4_1.Text) Or Not IsNumeric(Me.txtTest5_1.Text) Then GoTo NotNumber
                            cmd.CommandText = "UPDATE Specifications SET  Isolation = '" & Me.txtTest3_1.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  AmplitudeBalance = '" & Me.txtTest4_1.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            If ck1Expanded.CheckState = CheckState.Checked Then
                                cmd.CommandText = "UPDATE Specifications SET  AB_tf = 1" & Expression
                                cmd.ExecuteNonQuery()
                                cmd.CommandText = "UPDATE Specifications SET  AB_ex = '" & txtTest4_1_exp.Text & "'" & Expression
                                cmd.ExecuteNonQuery()
                                cmd.CommandText = "UPDATE Specifications SET  AB_start1 = '" & tx11Start.Text & "'" & Expression
                                cmd.ExecuteNonQuery()
                                cmd.CommandText = "UPDATE Specifications SET  AB_stop1 = '" & tx11Stop.Text & "'" & Expression
                                cmd.ExecuteNonQuery()
                                cmd.CommandText = "UPDATE Specifications SET  AB_start2 = '" & tx12Start.Text & "'" & Expression
                                cmd.ExecuteNonQuery()
                                cmd.CommandText = "UPDATE Specifications SET  AB_stop2 = '" & tx12Stop.Text & "'" & Expression
                                cmd.ExecuteNonQuery()
                            Else
                                cmd.CommandText = "UPDATE Specifications SET  AB_tf = 0" & Expression
                                cmd.ExecuteNonQuery()
                            End If
                            cmd.CommandText = "UPDATE Specifications SET  PhaseBalance = '" & Me.txtTest5_1.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                        End If
                        If Me.txtTest1_1.Text = "" Or Me.txtTest2_1.Text = "" Or Me.cmbJob.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtTest1_1.Text) Or Not IsNumeric(Me.txtTest2_1.Text) Then GoTo NotNumber

                        If Me.cmbPart.Text = "" Or Me.txtTitle.Text = "" Or Me.txtStartFreq_1.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStartFreq_1.Text) Then GoTo NotNumber

                        If Me.txtStopFreq_1.Text = "" Or Me.txtPower_1.Text = "" Or Me.txtQuantity.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStopFreq_1.Text) Or Not IsNumeric(Me.txtPower_1.Text) Or Not IsNumeric(Me.txtQuantity.Text) Then GoTo NotNumber

                        If Me.cmbSpecType.Text = "" Or Me.txtPower_1.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtPower_1.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        'cmd.CommandText = "UPDATE Specifications SET  Isolation2 = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  InsertionLoss = '" & Me.txtTest1_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  VSWR = '" & Me.txtTest2_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  JobNumber = '" & Me.cmbJob.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PartNumber = '" & FormatPartNum() & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Title = '" & Me.txtTitle.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Quantity = '" & Me.txtQuantity.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PPH = " & Me.txtPPH.Text & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StartFreqMHz = '" & Me.txtStartFreq_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StopFreqMHz = '" & Me.txtStopFreq_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Power = '" & Me.txtPower_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  SpecType = '" & Me.cmbSpecType.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Bypass = '" & M_bypass & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  FailPercent = '" & Me.txtFail.Text & "'" & Expression
                        cmd.ExecuteNonQuery()

                        'cmd.CommandText = "UPDATE Specifications SET  OutputPortNumber = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                        'cmd.CommandText = "UPDATE Specifications SET  CutOffFreqMHz = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                    Case 1
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            'If Me.txtTest3_2.Text = "" Or Me.txtTest4_2.Text = "" Or Me.txtTest5_2.Text = "" Then GoTo StupidUser
                            ' If Not IsNumeric(Me.txtTest3_2.Text) Or Not IsNumeric(Me.txtTest4_2.Text) Or Not IsNumeric(Me.txtTest5_2.Text) Then GoTo NotNumber
                            'cmd.CommandText = "UPDATE Specifications SET  Isolation = '" & Me.txtTest3_2.Text & "'" & Expression
                            'cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  AmplitudeBalance = '" & Me.txtTest4_2.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  PhaseBalance = '" & Me.txtTest5_2.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                        End If
                        If Me.txtTest1_2.Text = "" Or Me.txtTest2_2.Text = "" Or Me.cmbJob.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtTest1_2.Text) Or Not IsNumeric(Me.txtTest2_2.Text) Then GoTo NotNumber

                        If Me.cmbPart.Text = "" Or Me.txtTitle.Text = "" Or Me.txtStartFreq_2.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStartFreq_2.Text) Then GoTo NotNumber

                        If Me.txtStopFreq_2.Text = "" Or Me.txtPower_2.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStopFreq_2.Text) Or Not IsNumeric(Me.txtPower_2.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        If Me.cmbSpecType.Text = "" Or Me.txtPower_2.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtPower_2.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        'cmd.CommandText = "UPDATE Specifications SET  Isolation2 = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  InsertionLoss = '" & Me.txtTest1_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  VSWR = '" & Me.txtTest2_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  JobNumber = '" & Me.cmbJob.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PartNumber = '" & FormatPartNum() & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Title = '" & Me.txtTitle.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Quantity = '" & Me.txtQuantity.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PPH = " & Me.txtPPH.Text & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StartFreqMHz = '" & Me.txtStartFreq_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StopFreqMHz = '" & Me.txtStopFreq_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Power = '" & Me.txtPower_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  SpecType = '" & Me.cmbSpecType.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Bypass = '" & M_bypass & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  FailPercent = '" & Me.txtFail.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        'cmd.CommandText = "UPDATE Specifications SET  OutputPortNumber = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                        'cmd.CommandText = "UPDATE Specifications SET  CutOffFreqMHz = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                    Case 2
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Me.txtTest3_3.Text = "" Or Me.txtTest4_3.Text = "" Or Me.txtTest5_3.Text = "" Then GoTo StupidUser
                            If Not IsNumeric(Me.txtTest3_3.Text) Or Not IsNumeric(Me.txtTest4_3.Text) Or Not IsNumeric(Me.txtTest5_3.Text) Then GoTo NotNumber
                            cmd.CommandText = "UPDATE Specifications SET  Isolation = '" & Me.txtTest3_3.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  AmplitudeBalance = '" & Me.txtTest4_3.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  PhaseBalance = '" & Me.txtTest5_3.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                        End If
                        If Me.txtTest1_3.Text = "" Or Me.txtTest2_3.Text = "" Or Me.cmbJob.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtTest1_3.Text) Or Not IsNumeric(Me.txtTest2_3.Text) Then GoTo NotNumber

                        If Me.cmbPart.Text = "" Or Me.txtTitle.Text = "" Or Me.txtStartFreq_3.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStartFreq_3.Text) Then GoTo NotNumber

                        If Me.txtStopFreq_3.Text = "" Or Me.txtPower_3.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStopFreq_3.Text) Or Not IsNumeric(Me.txtPower_3.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        If Me.cmbSpecType.Text = "" Or Me.txtPower_3.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtPower_3.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        'cmd.CommandText = "UPDATE Specifications SET  Isolation2 = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  InsertionLoss = '" & Me.txtTest1_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  VSWR = '" & Me.txtTest2_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  JobNumber = '" & Me.cmbJob.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PartNumber = '" & FormatPartNum() & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Title = '" & Me.txtTitle.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Quantity = '" & Me.txtQuantity.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PPH = " & Me.txtPPH.Text & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StartFreqMHz = '" & Me.txtStartFreq_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StopFreqMHz = '" & Me.txtStopFreq_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Power = '" & Me.txtPower_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  SpecType = '" & Me.cmbSpecType.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Coupling = '" & Me.txtTest3_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Directivity = '" & Me.txtTest4_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        
                        cmd.CommandText = "UPDATE Specifications SET  CoupledFlatness = '" & Me.txtTest5_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Bypass = '" & M_bypass & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  FailPercent = '" & Me.txtFail.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        If ckPorts.CheckState = CheckState.Checked Then
                            cmd.CommandText = "UPDATE Specifications SET  SwPorts = 1" & Expression
                            cmd.ExecuteNonQuery()
                        Else
                            cmd.CommandText = "UPDATE Specifications SET  SwPorts = 0" & Expression
                            cmd.ExecuteNonQuery()
                        End If
                        If ckPlusMinus.CheckState = CheckState.Checked Then
                            cmd.CommandText = "UPDATE Specifications SET  COUPPlus = '" & Me.COUPPlus.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  COUPMinus = '" & Me.COUPMinus.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  COUP_DualSpec = 1" & Expression
                            cmd.ExecuteNonQuery()
                        Else
                            cmd.CommandText = "UPDATE Specifications SET  COUPPlusMinus = '" & Me.COUPPlusMinus.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  COUP_DualSpec = 0" & Expression
                            cmd.ExecuteNonQuery()
                        End If

                        'cmd.CommandText = "UPDATE Specifications SET  OutputPortNumber = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                        'cmd.CommandText = "UPDATE Specifications SET  CutOffFreqMHz = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                    Case 3
                        If Me.txtTest1_4.Text = "" Or Me.txtTest2_4.Text = "" Or Me.cmbJob.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtTest1_4.Text) Or Not IsNumeric(Me.txtTest2_4.Text) Then GoTo NotNumber

                        If Me.txtTest4_4.Text = "" Or Me.txtTest5_4.Text = "" Or Me.txtTest3_4.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtTest4_4.Text) Or Not IsNumeric(Me.txtTest5_4.Text) Or Not IsNumeric(Me.txtTest3_4.Text) Then GoTo NotNumber

                        If Me.cmbPart.Text = "" Or Me.txtTitle.Text = "" Or Me.txtStartFreq_4.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStartFreq_4.Text) Then GoTo NotNumber

                        If Me.txtStopFreq_4.Text = "" Or Me.txtPower_4.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStopFreq_4.Text) Or Not IsNumeric(Me.txtPower_4.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        If Me.cmbSpecType.Text = "" Or Me.txtPower_4.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtPower_4.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        cmd.CommandText = "UPDATE Specifications SET  Isolation = '" & Me.txtTest3_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  InsertionLoss = '" & Me.txtTest1_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  AmplitudeBalance = '" & Me.txtTest4_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PhaseBalance = '" & Me.txtTest5_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  VSWR = '" & Me.txtTest2_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  JobNumber = '" & Me.cmbJob.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PartNumber = '" & FormatPartNum() & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Title = '" & Me.txtTitle.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Quantity = '" & Me.txtQuantity.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PPH = " & Me.txtPPH.Text & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StartFreqMHz = '" & Me.txtStartFreq_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StopFreqMHz = '" & Me.txtStopFreq_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Power = '" & Me.txtPower_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  SpecType = '" & Me.cmbSpecType.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Bypass = '" & M_bypass & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  FailPercent = '" & Me.txtFail.Text & "'" & Expression
                        cmd.ExecuteNonQuery()

                        If Me.txtOutputPorts.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtOutputPorts.Text) Then GoTo NotNumber
                        cmd.CommandText = "UPDATE Specifications SET  OutputPortNumber = '" & Me.txtOutputPorts.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        If txtISOFreq_2.Text <> "" And txtISOFreq_2.Text <> "N/A" Then
                            cmd.CommandText = "UPDATE Specifications SET  CutOffFreqMHz = '" & txtISOFreq_2.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                        End If
                        If Me.txtTest3_5.Text <> "" Then
                            cmd.CommandText = "UPDATE Specifications SET  Isolation2 = '" & Me.txtTest3_5.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                        End If
                    Case 4
                        If Me.txtTest1_5.Text = "" Or Me.txtTest2_5.Text = "" Or Me.cmbJob.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtTest1_5.Text) Or Not IsNumeric(Me.txtTest2_5.Text) Then GoTo NotNumber

                        cmd.CommandText = "UPDATE Specifications SET  InsertionLoss = '" & Me.txtTest1_5.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  VSWR = '" & Me.txtTest2_5.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        If ilchk.CheckState = CheckState.Checked Then
                            If Me.txtTest1_5_exp.Text = "" Or Me.txtTest2_5.Text = "" Or Me.txtILFreq_1.Text = "" Or Me.txtILFreq_2.Text = "" Or Me.txtILFreq_3.Text = "" Or Me.txtILFreq_4.Text = "" Then GoTo StupidUser
                            If Not IsNumeric(Me.txtTest1_5_exp.Text) Or Not IsNumeric(Me.txtILFreq_1.Text) Or Not IsNumeric(Me.txtILFreq_2.Text) Or Not IsNumeric(Me.txtILFreq_3.Text) Or Not IsNumeric(Me.txtILFreq_4.Text) Then GoTo NotNumber

                            cmd.CommandText = "UPDATE Specifications SET  IL_exp_tf = 1" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  IL_ex = '" & txtTest1_5_exp.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  IL_start1 = '" & txtILFreq_1.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  IL_stop1 = '" & txtILFreq_2.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  IL_start2 = '" & txtILFreq_3.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  IL_stop2 = '" & txtILFreq_4.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                        Else
                            cmd.CommandText = "UPDATE Specifications SET  IL_exp_tf = 0" & Expression
                            cmd.ExecuteNonQuery()
                        End If

                        cmd.CommandText = "UPDATE Specifications Set Test1 = 1" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications Set Test2 = 1" & Expression
                        cmd.ExecuteNonQuery()

                        cmd.CommandText = "UPDATE Specifications SET  JobNumber = '" & Me.cmbJob.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PartNumber = '" & FormatPartNum() & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Title = '" & Me.txtTitle.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Quantity = '" & Me.txtQuantity.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PPH = " & Me.txtPPH.Text & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StartFreqMHz = '" & Me.txtStartFreq_5.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StopFreqMHz = '" & Me.txtStopFreq_5.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Power = '" & Me.txtPower_5.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  SpecType = '" & Me.cmbSpecType.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Bypass = '" & M_bypass & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  FailPercent = '" & Me.txtFail.Text & "'" & Expression
                        cmd.ExecuteNonQuery()

                End Select
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                cmd.Connection = atsLocal
                Select Case Index
                    Case 0
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Me.txtTest3_1.Text = "" Or Me.txtTest4_1.Text = "" Or Me.txtTest5_1.Text = "" Then GoTo StupidUser
                            If Not IsNumeric(Me.txtTest3_1.Text) Or Not IsNumeric(Me.txtTest4_1.Text) Or Not IsNumeric(Me.txtTest5_1.Text) Then GoTo NotNumber
                            cmd.CommandText = "UPDATE Specifications SET  Iso = '" & Me.txtTest3_1.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  AmplitudeBalance = '" & Me.txtTest4_1.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  PhaseBalance = '" & Me.txtTest5_1.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                        End If
                        If Me.txtTest1_1.Text = "" Or Me.txtTest2_1.Text = "" Or Me.cmbJob.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtTest1_1.Text) Or Not IsNumeric(Me.txtTest2_1.Text) Then GoTo NotNumber

                        If Me.cmbPart.Text = "" Or Me.txtTitle.Text = "" Or Me.txtStartFreq_1.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStartFreq_1.Text) Then GoTo NotNumber

                        If Me.txtStopFreq_1.Text = "" Or Me.txtPower_1.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStopFreq_1.Text) Or Not IsNumeric(Me.txtPower_1.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        If Me.cmbSpecType.Text = "" Or Me.txtPower_1.Text = "" Or Me.txtQuantity.Text = "" Or txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtPower_1.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        'cmd.CommandText = "UPDATE Specifications SET  Isolation2 = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  InsertionLoss = '" & Me.txtTest1_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  VSWR = '" & Me.txtTest2_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  JobNumber = '" & Me.cmbJob.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PartNumber = '" & FormatPartNum() & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Title = '" & Me.txtTitle.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Quantity = '" & Me.txtQuantity.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PPH = " & Me.txtPPH.Text & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StartFreqMHz = '" & Me.txtStartFreq_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StopFreqMHz = '" & Me.txtStopFreq_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Power = '" & Me.txtPower_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  SpecType = '" & Me.cmbSpecType.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Bypass = '" & M_bypass & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  FailPercent = '" & Me.txtFail.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        'cmd.CommandText = "UPDATE Specifications SET  OutputPortNumber = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                        'cmd.CommandText = "UPDATE Specifications SET  CutOffFreqMHz = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                    Case 1
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            'If Me.txtTest3_2.Text = "" Or Me.txtTest4_2.Text = "" Or Me.txtTest5_2.Text = "" Then GoTo StupidUser
                            ' If Not IsNumeric(Me.txtTest3_2.Text) Or Not IsNumeric(Me.txtTest4_2.Text) Or Not IsNumeric(Me.txtTest5_2.Text) Then GoTo NotNumber
                            'cmd.CommandText = "UPDATE Specifications SET  Iso = '" & Me.txtTest3_2.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  AmplitudeBalance = '" & Me.txtTest4_2.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  PhaseBalance = '" & Me.txtTest5_2.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                        End If
                        If Me.txtTest1_2.Text = "" Or Me.txtTest2_2.Text = "" Or Me.cmbJob.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtTest1_2.Text) Or Not IsNumeric(Me.txtTest2_2.Text) Then GoTo NotNumber

                        If Me.cmbPart.Text = "" Or Me.txtTitle.Text = "" Or Me.txtStartFreq_2.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStartFreq_2.Text) Then GoTo NotNumber

                        If Me.txtStopFreq_2.Text = "" Or Me.txtPower_2.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStopFreq_2.Text) Or Not IsNumeric(Me.txtPower_2.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        If Me.cmbSpecType.Text = "" Or Me.txtPower_2.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtPower_2.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        'cmd.CommandText = "UPDATE Specifications SET  Isolation2 = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  InsertionLoss = '" & Me.txtTest1_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  VSWR = '" & Me.txtTest2_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  JobNumber = '" & Me.cmbJob.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PartNumber = '" & FormatPartNum() & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Title = '" & Me.txtTitle.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Quantity = '" & Me.txtQuantity.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PPH = " & Me.txtPPH.Text & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StartFreqMHz = '" & Me.txtStartFreq_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StopFreqMHz = '" & Me.txtStopFreq_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Power = '" & Me.txtPower_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  SpecType = '" & Me.cmbSpecType.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Bypass = '" & M_bypass & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  FailPercent = '" & Me.txtFail.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        'cmd.CommandText = "UPDATE Specifications SET  OutputPortNumber = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                        'cmd.CommandText = "UPDATE Specifications SET  CutOffFreqMHz = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                    Case 2
                        If Index = 0 Or Index = 1 Or Index = 3 Then
                            If Me.txtTest3_3.Text = "" Or Me.txtTest4_3.Text = "" Or Me.txtTest5_3.Text = "" Then GoTo StupidUser
                            If Not IsNumeric(Me.txtTest3_3.Text) Or Not IsNumeric(Me.txtTest4_3.Text) Or Not IsNumeric(Me.txtTest5_3.Text) Then GoTo NotNumber
                            cmd.CommandText = "UPDATE Specifications SET  Iso = '" & Me.txtTest3_3.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  AmplitudeBalance = '" & Me.txtTest4_3.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                            cmd.CommandText = "UPDATE Specifications SET  PhaseBalance = '" & Me.txtTest5_3.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                        End If
                        If Me.txtTest1_3.Text = "" Or Me.txtTest2_3.Text = "" Or Me.cmbJob.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtTest1_3.Text) Or Not IsNumeric(Me.txtTest2_3.Text) Then GoTo NotNumber

                        If Me.cmbPart.Text = "" Or Me.txtTitle.Text = "" Or Me.txtStartFreq_3.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStartFreq_3.Text) Then GoTo NotNumber

                        If Me.txtStopFreq_3.Text = "" Or Me.txtPower_3.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStopFreq_3.Text) Or Not IsNumeric(Me.txtPower_3.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        If Me.cmbSpecType.Text = "" Or Me.txtPower_3.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtPower_3.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        'cmd.CommandText = "UPDATE Specifications SET  Isolation2 = '" & DBNull & "'" & Expression
                        'cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  InsertionLoss = '" & Me.txtTest1_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  VSWR = '" & Me.txtTest2_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  JobNumber = '" & Me.cmbJob.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PartNumber = '" & FormatPartNum() & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Title = '" & Me.txtTitle.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Quantity = '" & Me.txtQuantity.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PPH = " & Me.txtPPH.Text & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StartFreqMHz = '" & Me.txtStartFreq_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StopFreqMHz = '" & Me.txtStopFreq_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Power = '" & Me.txtPower_3.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  SpecType = '" & Me.cmbSpecType.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Bypass = '" & M_bypass & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  FailPercent = '" & Me.txtFail.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                    Case 3
                        If Me.txtTest1_4.Text = "" Or Me.txtTest2_4.Text = "" Or Me.cmbJob.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtTest1_4.Text) Or Not IsNumeric(Me.txtTest2_4.Text) Then GoTo NotNumber

                        If Me.cmbPart.Text = "" Or Me.txtTitle.Text = "" Or Me.txtStartFreq_4.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStartFreq_4.Text) Then GoTo NotNumber

                        If Me.txtStopFreq_4.Text = "" Or Me.txtPower_4.Text = "" Or Me.txtQuantity.Text = "" Or Me.txtPPH.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtStopFreq_4.Text) Or Not IsNumeric(Me.txtPower_4.Text) Or Not IsNumeric(Me.txtQuantity.Text) Or Not IsNumeric(Me.txtPPH.Text) Then GoTo NotNumber

                        If Me.cmbSpecType.Text = "" Or Me.txtPower_4.Text = "" Or Me.txtQuantity.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtPower_4.Text) Or Not IsNumeric(Me.txtQuantity.Text) Then GoTo NotNumber

                        cmd.CommandText = "UPDATE Specifications SET  Isolation = '" & Me.txtTest3_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()

                        cmd.CommandText = "UPDATE Specifications SET  InsertionLoss = '" & Me.txtTest1_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  VSWR = '" & Me.txtTest2_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  JobNumber = '" & Me.cmbJob.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PartNumber = '" & FormatPartNum() & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Title = '" & Me.txtTitle.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Quantity = '" & Me.txtQuantity.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PPH = " & Me.txtPPH.Text & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StartFreqMHz = '" & Me.txtStartFreq_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StopFreqMHz = '" & Me.txtStopFreq_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Power = '" & Me.txtPower_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  SpecType = '" & Me.cmbSpecType.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Bypass = '" & M_bypass & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  FailPercent = '" & Me.txtFail.Text & "'" & Expression
                        cmd.ExecuteNonQuery()

                        If Me.txtOutputPorts.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtOutputPorts.Text) Then GoTo NotNumber
                        cmd.CommandText = "UPDATE Specifications SET  OutputPortNumber = '" & Me.txtOutputPorts.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        If txtISOFreq_2.Text <> "0" Then
                            cmd.CommandText = "UPDATE Specifications SET  CutOffFreqMHz = '" & txtISOFreq_2.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                        End If
                        If Me.txtTest3_5.Text <> "" Then
                            cmd.CommandText = "UPDATE Specifications SET  Isolation2 = '" & Me.txtTest3_5.Text & "'" & Expression
                            cmd.ExecuteNonQuery()
                        End If

                    Case 4
                        If Me.txtTest1_5.Text = "" Or Me.txtTest2_5.Text = "" Or Me.cmbJob.Text = "" Then GoTo StupidUser
                        If Not IsNumeric(Me.txtTest1_5.Text) Or Not IsNumeric(Me.txtTest2_5.Text) Then GoTo NotNumber

                        cmd.CommandText = "UPDATE Specifications SET  InsertionLoss = '" & Me.txtTest1_5.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  VSWR = '" & Me.txtTest2_5.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  JobNumber = '" & Me.cmbJob.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PartNumber = '" & FormatPartNum() & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Title = '" & Me.txtTitle.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Quantity = '" & Me.txtQuantity.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  PPH = " & Me.txtPPH.Text & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StartFreqMHz = '" & Me.txtStartFreq_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  StopFreqMHz = '" & Me.txtStopFreq_4.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Power = '" & Me.txtPower_5.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  SpecType = '" & Me.cmbSpecType.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  Bypass = '" & M_bypass & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE Specifications SET  FailPercent = '" & Me.txtFail.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                End Select
                atsLocal.Close()
            End If
            '"~~>************************SQL Or Access Conection****************************



            '"~~>************************Port Config**********************
            If Index <> 3 Then

                Expression = " where PartNumber = '" & FormatPartNum() & "' And JobNumber = '" & Me.cmbJob.Text & "'"
                SQLstr = "SELECT * from PortConfig" & Expression

                If SQL.CheckforRow(SQLstr, "LocalSpecs") <= 0 Then
                    SQLstr = "Insert Into PortConfig (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
                    SQL.ExecuteSQLCommand(SQLstr, "LocalSpecs")
                End If

                '"~~>************************SQL Or Access Conection****************************
                If SQLAccess Then
                    Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                    Dim cmd As SqlCommand = New SqlCommand()
                    ats.Open()
                    cmd.Connection = ats
                    '"~~>************************SQL Or Access Conection****************************
                    If Index = 1 Then
                        cmd.CommandText = "UPDATE PortConfig SET  J1J1 = '" & txtJ1J1_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J1J2 = '" & txtJ1J2_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J1J3 = '" & txtJ1J3_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                    Else
                        cmd.CommandText = "UPDATE PortConfig SET  J1J1 = '" & txtJ1J1_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J1J2 = '" & txtJ1J2_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J1J3 = '" & txtJ1J3_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J1J4 = '" & txtJ1J4_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J2J1 = '" & txtJ2J1_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J2J2 = '" & txtJ2J2_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J2J3 = '" & txtJ2J3_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J2J4 = '" & txtJ2J4_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J3J1 = '" & txtJ3J1_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J3J2 = '" & txtJ3J2_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J3J3 = '" & txtJ3J3_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J3J4 = '" & txtJ3J4_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J4J1 = '" & txtJ4J1_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J4J2 = '" & txtJ4J2_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J4J3 = '" & txtJ4J3_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J4J4 = '" & txtJ4J4_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                    End If
                    ats.Close()
                Else
                    Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
                    Dim atsLocal As New OleDb.OleDbConnection
                    atsLocal.ConnectionString = strConnectionString
                    Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                    atsLocal.Open()
                    cmd.Connection = atsLocal
                    '"~~>************************SQL Or Access Conection****************************
                    If Index = 1 Then
                        cmd.CommandText = "UPDATE PortConfig SET  J1J1 = '" & txtJ1J1_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J1J2 = '" & txtJ1J2_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J1J3 = '" & txtJ1J3_2.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                    Else
                        cmd.CommandText = "UPDATE PortConfig SET  J1J1 = '" & txtJ1J1_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J1J2 = '" & txtJ1J2_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J1J3 = '" & txtJ1J3_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J1J4 = '" & txtJ1J4_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J2J1 = '" & txtJ2J1_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J2J2 = '" & txtJ2J2_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J2J3 = '" & txtJ2J3_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J2J4 = '" & txtJ2J4_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J3J1 = '" & txtJ3J1_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J3J2 = '" & txtJ3J2_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J3J3 = '" & txtJ3J3_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J3J4 = '" & txtJ3J4_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J4J1 = '" & txtJ4J1_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J4J2 = '" & txtJ4J2_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J4J3 = '" & txtJ4J3_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                        cmd.CommandText = "UPDATE PortConfig SET  J4J4 = '" & txtJ4J4_1.Text & "'" & Expression
                        cmd.ExecuteNonQuery()
                    End If
                    atsLocal.Close()
                End If
            End If

            Dim Count As Integer = 0
            frmAUTOTEST.cmbJob.Items.Clear()
            frmAUTOTEST.cmbJob.Items.Add(" ")
            frmAUTOTEST.cmbJob.Items.Add("Add New")
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand("SELECT DISTINCT JobNumber from Specifications", ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While dr.Read
                    frmAUTOTEST.cmbJob.Items.Add(CType(dr.GetValue(0), String))
                    Count = Count + 1
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT DISTINCT JobNumber from Specifications", atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    frmAUTOTEST.cmbJob.Items.Add(CType(drLocal.GetValue(0), String))
                    Count = Count + 1
                End While
                atsLocal.Close()
            End If
            


            Me.Hide()
            If ActivePage = "Full Auto" Then
                frmAUTOTEST.Show()
                frmAUTOTEST.cmbJob.Text = Job
                frmAUTOTEST.cmbPart.Text = Part
                frmAUTOTEST.txtStartFreq.Text = SpecStartFreq
                frmAUTOTEST.txtStopFreq.Text = SpecStopFreq
            Else
                frmMANUALTEST.Show()
                frmMANUALTEST.cmbJob.Text = Job
                frmMANUALTEST.cmbPart.Text = Part
                frmMANUALTEST.txtStartFreq.Text = SpecStartFreq
                frmMANUALTEST.txtStopFreq.Text = SpecStopFreq
            End If
            Exit Sub
StupidUser:
            MYMsgBox("Please fill out all required fields")
            Exit Sub
NoRatio:
            MYMsgBox("Ratio is already included in VSWR Spec. Please remove the :")
NotNumber:
            MYMsgBox("Please recheck out all required fields. You must have all Numerials and no Characters")
            Exit Sub
        Catch ex As Exception
            Me.Hide()
            If ActivePage = "Full Auto" Then
                frmAUTOTEST.Show()
            Else
                frmMANUALTEST.Show()
            End If
        End Try
    End Sub

    Private Sub txtISOFreq_1_TextChanged(sender As Object, e As EventArgs)
        txtISOFreq_3 = txtISOFreq_2
    End Sub

    Private Sub txtISOFreq_2_TextChanged(sender As Object, e As EventArgs)
        txtISOFreq_3 = txtISOFreq_2
    End Sub

    Private Sub txtISOFreq_3_TextChanged(sender As Object, e As EventArgs)
        txtISOFreq_3 = txtISOFreq_2
    End Sub

    Private Sub txtISOFreq_4_TextChanged(sender As Object, e As EventArgs)
        txtISOFreq_3 = txtISOFreq_2
    End Sub


    Private Sub txtStartFreq_1_TextChanged(sender As Object, e As EventArgs)
        txtISOFreq_1 = txtStartFreq_4
    End Sub
    Private Sub txtStartFreq_2_TextChanged(sender As Object, e As EventArgs) Handles txtStartFreq_2.TextChanged
        txtISOFreq_1 = txtStartFreq_4
    End Sub
    Private Sub txtStartFreq_3_TextChanged(sender As Object, e As EventArgs)
        txtISOFreq_1 = txtStartFreq_4
    End Sub
    Private Sub txtStartFreq_4_TextChanged(sender As Object, e As EventArgs)
        txtISOFreq_1 = txtStartFreq_4
    End Sub

    Private Sub txtStopFreq_1_TextChanged(sender As Object, e As EventArgs)
        txtISOFreq_4 = txtStopFreq_4
    End Sub

    Private Sub txtStopFreq_2_TextChanged(sender As Object, e As EventArgs) Handles txtStopFreq_2.TextChanged
        txtISOFreq_4 = txtStopFreq_4
    End Sub

    Private Sub txtStopFreq_3_TextChanged(sender As Object, e As EventArgs)
        txtISOFreq_4 = txtStopFreq_4
    End Sub

    Private Sub txtStopFreq_4_TextChanged(sender As Object, e As EventArgs)
        txtISOFreq_4 = txtStopFreq_4
    End Sub

    Private Function FormatPartNum() As String
        Dim PTArray
        PTArray = Split(Me.cmbPart.Text, "-")
        FormatPartNum = Trim(PTArray(0)) & "-" & Trim(PTArray(1))
    End Function

    Private Sub sup_bypass_CheckedChanged(sender As Object, e As EventArgs) Handles sup_bypass.CheckedChanged
        Master_bypass = False
        Dim bypasstemp = sup_bypass.CheckState

        If sup_bypass.CheckState = CheckState.Checked And Not byp Then
            Dim SPEC As New Supervisor
            SPEC.StartPosition = FormStartPosition.Manual
            ' SPEC.Location = New Point(Globals.XLocation + NewXLocation, Globals.YLocation + Globals.XSize)
            SPEC.ShowDialog()
            Percent_bypass = False
        End If

        If Not Master_bypass And Not byp Then
            sup_bypass.CheckState = CheckState.Unchecked
        End If
    End Sub

    Private Sub txtFail_TextChanged(sender As Object, e As EventArgs) Handles txtFail.TextChanged
        Dim failTemp = txtFail.Text
        If Not IsNumeric(txtFail.Text) Then
            MYMsgBox("Enter Numbers only", , "No Characters")
            txtFail.Text = GetFailPercent()
            Me.Close()
        End If
        If Not Percent_bypass Then
            Dim SPEC As New Supervisor
            SPEC.StartPosition = FormStartPosition.Manual
            ' SPEC.Location = New Point(Globals.XLocation + NewXLocation, Globals.YLocation + Globals.XSize)
            SPEC.ShowDialog()
        End If

        If Not Percent_bypass Then
            txtFail.Text = GetFailPercent()
        End If
    End Sub
    Private Sub ck1Expanded_CheckedChanged(sender As Object, e As EventArgs) Handles ck1Expanded.CheckedChanged
        If ck1Expanded.CheckState = CheckState.Checked Then
            lbl1Start.Visible = True
            lbl1Stop.Visible = True
            tx11Start.Visible = True
            tx11Stop.Visible = True
            tx12Start.Visible = True
            tx12Stop.Visible = True
            txtTest4_1_exp.Enabled = True
            If Not SpecAB_TF Then
                tx11Start.Text = txtStartFreq_1.Text
                tx12Stop.Text = txtStopFreq_1.Text
                txtTest4_1_exp.Text = txtTest4_1.Text
            End If

        Else
            lbl1Start.Visible = False
            lbl1Stop.Visible = False
            tx11Start.Visible = False
            tx11Stop.Visible = False
            tx12Start.Visible = False
            tx12Stop.Visible = False
            txtTest4_1_exp.Enabled = False

        End If

    End Sub


    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.CheckState = CheckState.Checked Then
            txtISOFreq_1.Visible = True
            txtISOFreq_2.Visible = True
            txtISOFreq_3.Visible = True
            txtISOFreq_4.Visible = True
            Mhz1.Visible = True
            Mhz2.Visible = True
            txtTest3_5.Enabled = True
            dB2.Enabled = True
            ISO_TF = True
            If Not SpecAB_TF Then
                txtISOFreq_1.Text = txtStartFreq_4.Text
                txtISOFreq_4.Text = txtStopFreq_4.Text
            End If

        Else
            txtISOFreq_1.Visible = False
            txtISOFreq_2.Visible = False
            txtISOFreq_3.Visible = False
            txtISOFreq_4.Visible = False
            Mhz1.Visible = False
            Mhz2.Visible = False
            txtISOFreq_1.Text = 0
            txtISOFreq_2.Text = 0
            txtISOFreq_3.Text = 0
            txtISOFreq_4.Text = 0
            txtTest3_5.Text = ""
            txtTest3_5.Enabled = False
            dB2.Enabled = True
            ISO_TF = False
        End If
    End Sub

    Private Sub txtISOFreq_2_TextChanged_1(sender As Object, e As EventArgs) Handles txtISOFreq_2.TextChanged
        txtISOFreq_3.Text = txtISOFreq_2.Text
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles ilchk.CheckedChanged
        If ilchk.CheckState = CheckState.Checked Then
            txtILFreq_1.Visible = True
            txtILFreq_2.Visible = True
            txtILFreq_3.Visible = True
            txtILFreq_4.Visible = True
            MH11.Visible = True
            MH22.Visible = True
            txtTest1_5_exp.Enabled = True
            db22.Enabled = True
            IL_TF = True
            If Not IL_TF Then
                txtILFreq_1.Text = txtStartFreq_4.Text
                txtILFreq_4.Text = txtStopFreq_4.Text
            End If

        Else
            txtISOFreq_1.Visible = False
            txtISOFreq_2.Visible = False
            txtISOFreq_3.Visible = False
            txtISOFreq_4.Visible = False
            Mhz1.Visible = False
            Mhz2.Visible = False
            txtISOFreq_1.Text = 0
            txtISOFreq_2.Text = 0
            txtISOFreq_3.Text = 0
            txtISOFreq_4.Text = 0
            txtTest2_5.Text = ""
            txtTest1_5_exp.Enabled = False
            dB2.Enabled = True
            IL_TF = False
        End If
    End Sub

    Private Sub ckPorts_CheckedChanged(sender As Object, e As EventArgs) Handles ckPorts.CheckedChanged
        If ckPorts.CheckState = CheckState.Checked Then
            ckPorts.Text = "6 Port Fixture"
            SwitchPorts = 1
        Else
            ckPorts.Text = "4 Port Fixture"
            SwitchPorts = 0
        End If
    End Sub

   
    Private Sub ckPlusMinus_CheckedChanged(sender As Object, e As EventArgs) Handles ckPlusMinus.CheckedChanged
        If ckPlusMinus.CheckState = CheckState.Checked Then
            COupDualSpec = True
            COUPPlusMinus.Visible = False
            PlusMinus.Visible = False
            Plus.Visible = True
            Minus.Visible = True
            COUPPlus.Visible = True
            COUPMinus.Visible = True
        Else
            COupDualSpec = False
            COUPPlusMinus.Visible = True
            PlusMinus.Visible = True
            Plus.Visible = False
            Minus.Visible = False
            COUPPlus.Visible = False
            COUPMinus.Visible = False
        End If
    End Sub

End Class


