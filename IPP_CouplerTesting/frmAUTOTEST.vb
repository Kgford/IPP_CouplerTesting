Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection



Public Class frmAUTOTEST
    Private Incomplete As Boolean = False
    Private GlobalFailMax As Integer = 10
    Private retestFailMax As Integer = 5
    Private TestFailMax As Integer = 5
    Private GlobalFail As Integer
    Private GlobalFailed As Boolean = False
    Private Test1Fail As Integer
    Private Test2Fail As Integer
    Private TEST3Fail As Integer
    Private TEST4Fail As Integer
    Private TEST4LFail As Integer
    Private TEST4HFail As Integer
    Private TEST5Fail As Integer
    Private Retest1Fail As Integer
    Private Retest2Fail As Integer
    Private RetEST3Fail As Integer
    Private RetEST4Fail As Integer
    Private RetEST5Fail As Integer
    Private Markertest1 As Boolean = False
    Private Markertest2 As Boolean = False
    Private Markertest3 As Boolean = False
    Private Markertest4 As Boolean = False
    Private Markertest5 As Boolean = False
    Private UUTFail As Integer
    Private ILFail As Integer
    Private LOTFail As Integer
    Private TEST1FailRetest As Integer
    Private TEST2FailRetest As Integer
    Private TEST3FailRetest As Integer
    Private TEST4FailRetest As Integer
    Private TEST5FailRetest As Integer
    Private TEST1PASS As Boolean = True
    Private TEST2PASS As Boolean = True
    Private TEST3PASS As Boolean = True
    Private TEST4PASS As Boolean = True
    Private TEST4LPASS As Boolean = True
    Private TEST4HPASS As Boolean = True
    Private TEST5PASS As Boolean = True
    Private SelectTest1 As Boolean = False
    Private SelectTest2 As Boolean = False
    Private SelectTest3 As Boolean = False
    Private SelectTest4 As Boolean = False
    Private SelectTest5 As Boolean = False
    Private FailCount As Integer
    Private LastUUT As Integer
    Private SpecIndex As Integer
    Private SpecType As String
    Private DontclickTheButton As Boolean
    Private RunningOffsets As Boolean
    Public Uploading As Boolean
    Private SpecTest1 As String
    Private SpecTest2 As String
    Private SpecTest3 As String
    Private SpecTest4 As String
    Private SpecTest4_exp As String
    Private SpecTest5 As String
    Private ILSetDone As Boolean
    Private TestPaused = False
    Private StartTime As Date
    Private StopTime As Date
    Private TestTime As TimeSpan

    Public Sub New()
        InitializeComponent()
        Try
            SQLAccess = My.Computer.Network.Ping("INN-SQLEXPRESS")
        Catch
            SQLAccess = False
        End Try
        SQLVerified = SQLAccess
        Try
            CleanUpEffeciency()
            TestRunningSignal(False)
            TestCompleteSignal(False)
            TestRunning = False
            NoInit = True
            NetworkAccess = CheckNetworkFolder()
            WorkStation = GetComputerName()
            If WorkStation = "mford" Then
                RFSwitch.Visible = True
                Simulation.Visible = True
            Else
                RFSwitch.Visible = False
                Simulation.Visible = False
                Pass.Visible = False
            End If
            Version = "Version: " & GetVersion()
            xTitle = "Freq MHz"
            yTitle = "Signal dBm"
            Temperature = "25"
            CalDue = Now
            SupervisorPassword = SQL.GetPassword()
            Dim SQLstr As String

            StatusLog.Items.Add("Inovative Power Products ATE")

            If NetworkAccess Then
                txtNet.Checked = True
            Else
                txtNet.Checked = False
            End If
            NetworkChecked()

            Me.txtVersion.Text = "Version: " + GetVersion()
            Uploading = False
            ILSetDone = False

            Me.EndLot.Enabled = False
            Me.EraseJob.Enabled = False

            Me.cmbJob.Items.Clear()
            Me.cmbJob.Text = " "
            Me.cmbJob.Items.Add("Add New")
            Me.cmbPart.Items.Clear()
            Me.cmbPart.Items.Add("Add New")

            Me.cmbPart.Items.Clear()
            SQLstr = "SELECT DISTINCT PartNumber from Specifications"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(0) = "IPP-" Then GoTo Skip1
                    Me.cmbPart.Items.Add(dr.Item(0))
                End While
Skip1:
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
                    If drLocal.Item(0) = "IPP-" Then GoTo Skip2
                    Me.cmbPart.Items.Add(drLocal.Item(0))
                End While
Skip2:
                atsLocal.Close()
            End If

            Me.cmbJob.Items.Clear()
            SQLstr = "SELECT DISTINCT JobNumber from Specifications"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(0) = "IPP-" Then GoTo Skip3
                    Me.cmbJob.Items.Add(dr.Item(0))
                End While
Skip3:
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
                    If drLocal.Item(0) = "IPP-" Then GoTo Skip4
                    Me.cmbJob.Items.Add(drLocal.Item(0))
                End While
Skip4:
                atsLocal.Close()
            End If

            Me.cmbVNA.Items.Clear()
            SQLstr = "SELECT * from DEVICES where DevType = 'VNA'"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    Me.cmbVNA.Items.Add(dr.Item(0))
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
                    Me.cmbVNA.Items.Add(drLocal.Item(0))
                End While
                atsLocal.Close()
            End If


            If GetVNAType() <> "" Then
                Me.cmbVNA.Text = GetVNAType()
            Else
                Me.cmbVNA.SelectedIndex = 2
            End If
            Dim SwNum As Integer
            Dim SerList As String = ""
            Dim Firm As Integer
            firstInitializeSwitch(SwNum, SerList, Firm)
            ExpectedProgress.Value = 0
            ActualProgress.Value = 0
            txtCurrentTime.Text = "READY"
        Catch
        End Try
    End Sub
    '*********************************************************************************************************************************************************
    '***********************************************************************FOLLOW THE LEADER*****************************************************************
    '*********************************************THIS CODE MAKES ALL CHILD FORMS FOLLOW THE MAIN FORM POSITION***********************************************
    '********************************************************THIS MUST BE UPDATED FOR ALL NEW FORMS*********************************************************** 
    '*********************************************************************************************************************************************************

    Private Sub Form1_Move(sender As Object, e As System.EventArgs) Handles MyBase.Move
        Dim frmCollection = System.Windows.Forms.Application.OpenForms

        globals.XLocation = Me.Location.X
        globals.YLocation = Me.Location.Y
        globals.XSize = Me.Size.Height
        globals.YSize = Me.Size.Width
        'If frmCollection.OfType(Of ModelConfig).Any Then
        '    frmCollection.Item(1).StartPosition = FormStartPosition.Manual
        '    frmCollection.Item(1).Location = New Point(XLocation, YLocation + XSize)
        'End If
    End Sub
    Public Function firstInitializeSwitch(retSwitchNumber As Integer, retSerialList As String, retFirmware As Integer) As String
        Dim ConnectionGood As Integer
        Try
            retSwitchNumber = GetNumberOfSwitches()
            retSerialList = GetSNlist()
            retFirmware = GetFirmware()
            Dim model = GetSwitchModel()


            ConnectionGood = SwitchCom.Connect ' Note requires a few milliseconds to connect
            System.Threading.Thread.Sleep(1000)

            firstInitializeSwitch = SwitchCom.Get24VConnection
            If firstInitializeSwitch.ToUpper.Contains("NOT CONNECTED") Then
                Me.RFSwitch.CheckState = CheckState.Unchecked
                SwitchedChecked = Me.RFSwitch.Checked
            End If
            SwitchModel = GetSwitchModel()
            If SwitchModel = "RC-1SP6T-A12" Then
                Me.cmbSwitch.Items.Clear()
                Me.cmbSwitch.Items.Add("Switch POS 1")
                Me.cmbSwitch.Items.Add("Switch POS 2")
                Me.cmbSwitch.Items.Add("Switch POS 3")
                Me.cmbSwitch.Items.Add("Switch POS 4")
                Me.cmbSwitch.Items.Add("Switch POS 5")
                Me.cmbSwitch.Items.Add("Switch POS 6")
                Me.cmbSwitch.Text = "Switch POS 1"
            Else
                Me.cmbSwitch.Items.Clear()
                Me.cmbSwitch.Items.Add("Switch POS 1")
                Me.cmbSwitch.Items.Add("Switch POS 2")
                Me.cmbSwitch.Items.Add("Switch POS 3")
                Me.cmbSwitch.Items.Add("Switch POS 4")
                Me.cmbSwitch.Text = "Switch POS 1"
            End If
        Catch ex As Exception

        End Try
    End Function
    Private Sub RetestSelect()
        Dim KeepTogether As Boolean

        '
        'Data2.Text = ""
        'Data3.Text = ""
        'Data4.Text = ""
        'Data5.Text = ""
        SelectTest1 = Me.ckTest1.Checked
        SelectTest2 = Me.ckTest2.Checked
        SelectTest3 = Me.ckTest3.Checked
        SelectTest4 = Me.ckTest4.Checked
        SelectTest5 = Me.ckTest5.Checked
        If SpecType <> "COMBINER/DIVIDER" And Not SpecType.Contains("DIRECTIONAL COUPLER") Then
            KeepTogether = True
        Else
            KeepTogether = False
        End If

        If TEST1PASS And Not KeepTogether Then
            Me.ckTest1.Checked = False
        Else
            If Not TEST1PASS Then
                If SelectTest1 Then Me.ckTest1.Checked = True
                Data1.Text = ""
            Else
                Me.ckTest1.Checked = False
            End If
        End If
        If TEST2PASS Then
            Me.ckTest2.Checked = False
        Else
            Data2.Text = ""
        End If

        If TEST3PASS Then
            Me.ckTest3.Checked = False
        Else
            Data3.Text = ""
        End If

        If TEST4PASS And Not KeepTogether Then
            Me.ckTest4.Checked = False
        ElseIf KeepTogether Then
            If ckTest1.Checked Then
                If SelectTest4 Then Me.ckTest4.Checked = True
                If SpecAB_TF Then
                    Data4L.Text = ""
                    Data4H.Text = ""
                Else
                    Data4.Text = ""
                End If

            Else
                Me.ckTest4.Checked = False
            End If
        Else
            Me.ckTest4.Checked = True
            Data4.Text = ""
        End If


        If TEST5PASS Then
            Me.ckTest5.Checked = False
        Else
            Data5.Text = ""
        End If

    End Sub

    'Private Sub EndLot_Click(sender As Object, e As EventArgs) Handles EndLot.CheckStateChanged
    '    'UploadJobData
    '    'FrmFlexGrid.ExeclData_click()
    '    ResetLot()
    '    cmbJob.Text = " "
    'End Sub
    Private Sub ckTest1_CheckedChanged(sender As Object, e As EventArgs) Handles ckTest1.CheckStateChanged
        Dim SQLstr As String
        Dim Expression As String
        Dim NonBool As Integer = 0
        If Not NoInit Then Exit Sub
        Try
            If SpecType <> "COMBINER/DIVIDER" And SpecType <> "SINGLE DIRECTIONAL COUPLER" And SpecType <> "DUAL DIRECTIONAL COUPLER" And SpecType <> "BI DIRECTIONAL COUPLER" Then
                If Me.ckTest4.Checked = True And Me.ckTest1.Checked = False Then Me.ckTest1.Checked = True
            End If
            Expression = " where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            SQLstr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            If SQL.CheckforRow(SQLstr, "NetworkSpecs") = 0 Then
                SQLstr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
                SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            End If

            NonBool = 0
            If Me.ckTest1.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test1 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest2.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test2 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest3.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test3 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest4.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test4 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest5.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test5 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")

        Catch ex As Exception

        End Try
    End Sub
    Private Sub ckTest2_CheckedChanged(sender As Object, e As EventArgs) Handles ckTest2.CheckStateChanged
        Dim SQLstr As String
        Dim Expression As String
        Dim NonBool As Integer = 0
        Try
            If Not NoInit Then Exit Sub
            If SpecType <> "COMBINER/DIVIDER" And SpecType <> "SINGLE DIRECTIONAL COUPLER" And SpecType <> "DUAL DIRECTIONAL COUPLER" And SpecType <> "BI DIRECTIONAL COUPLER" Then
                If Me.ckTest4.Checked = True And Me.ckTest1.Checked = False Then Me.ckTest1.Checked = True
            End If
            Expression = " where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            SQLstr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            If SQL.CheckforRow(SQLstr, "LocalSpecs") = 0 Then
                SQLstr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
                SQL.ExecuteSQLCommand(SQLstr, "LocalSpecs")
            End If

            NonBool = 0
            If Me.ckTest1.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test1 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest2.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test2 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest3.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test3 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest4.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test4 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest5.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test5 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")

        Catch ex As Exception

        End Try
    End Sub
    Private Sub ckTest3_CheckedChanged(sender As Object, e As EventArgs) Handles ckTest3.CheckStateChanged
        Dim SQLstr As String
        Dim Expression As String
        Dim NonBool As Integer = 0
        Try
            If Not NoInit Then Exit Sub
            If SpecType <> "COMBINER/DIVIDER" And SpecType <> "SINGLE DIRECTIONAL COUPLER" And SpecType <> "DUAL DIRECTIONAL COUPLER" And SpecType <> "BI DIRECTIONAL COUPLER" Then
                If Me.ckTest4.Checked = True And Me.ckTest1.Checked = False Then Me.ckTest1.Checked = True
            End If
            Expression = " where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            SQLstr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            If SQL.CheckforRow(SQLstr, "LocalSpecs") = 0 Then
                SQLstr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
                SQL.ExecuteSQLCommand(SQLstr, "LocalSpecs")
            End If
            NonBool = 0
            If Me.ckTest1.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test1 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest2.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test2 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest3.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test3 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest4.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test4 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest5.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test5 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")

        Catch ex As Exception

        End Try
    End Sub
    Private Sub ckTest4_CheckedChanged(sender As Object, e As EventArgs) Handles ckTest4.CheckStateChanged
        Dim SQLstr As String
        Dim Expression As String
        Dim NonBool As Integer = 0
        If Not NoInit Then Exit Sub
        Try
            If Not NoInit Then Exit Sub
            If SpecType <> "COMBINER/DIVIDER" And SpecType <> "SINGLE DIRECTIONAL COUPLER" And SpecType <> "DUAL DIRECTIONAL COUPLER" And SpecType <> "BI DIRECTIONAL COUPLER" Then
                If Me.ckTest4.Checked = True And Me.ckTest1.Checked = False Then Me.ckTest1.Checked = True
            End If
            Expression = " where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            SQLstr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            If SQL.CheckforRow(SQLstr, "NetworkSpecs") = 0 Then
                SQLstr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
                SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            End If
            NonBool = 0
            If Me.ckTest1.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test1 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest2.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test2 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest3.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test3 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest4.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test4 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest5.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test5 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")

        Catch ex As Exception

        End Try
    End Sub
    Private Sub ckTest5_CheckedChanged(sender As Object, e As EventArgs) Handles ckTest5.CheckStateChanged
        Dim SQLstr As String
        Dim Expression As String
        Dim NonBool As Integer = 0
        If Not NoInit Then Exit Sub
        Try

            If SpecType <> "COMBINER/DIVIDER" And SpecType <> "SINGLE DIRECTIONAL COUPLER" And SpecType <> "DUAL DIRECTIONAL COUPLER" And SpecType <> "BI DIRECTIONAL COUPLER" Then
                If Me.ckTest4.Checked = True And Me.ckTest1.Checked = False Then Me.ckTest1.Checked = True
            End If
            Expression = " where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            SQLstr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            If SQL.CheckforRow(SQLstr, "NetworkSpecs") = 0 Then
                SQLstr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
                SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            End If
            NonBool = 0
            If Me.ckTest1.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test1 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest2.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test2 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest3.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test3 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest4.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test4 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            NonBool = 0
            If Me.ckTest5.Checked Then NonBool = 1
            SQLstr = "UPDATE Specifications Set Test5 = " & NonBool & " " & Expression
            SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RetestReset()

        Me.ckTest1.Checked = SelectTest1
        Me.ckTest2.Checked = SelectTest2
        Me.ckTest3.Checked = SelectTest3
        Me.ckTest4.Checked = SelectTest4
        Me.ckTest5.Checked = SelectTest5

    End Sub

    Private Sub cmbJob_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbJob.SelectedIndexChanged
        Try
            If Me.cmbJob.GetItemText(Me.cmbJob.SelectedItem) <> " " And Uploading = False Then
                If Not DontclickTheButton Then
                    cmdStartTest.Enabled = False
                    Me.EraseTest.Enabled = False
                    If Not Simulation.Checked And Not Connected Then
                        ScanGPIB.connect("GPIB0::16::INSTR", GetTimeout())
                        Connected = True
                    End If
                    Dim SwNum As Integer
                    Dim SerList As String = ""
                    Dim Firm As Integer
                    Tests.InitializeSwitch(SwNum, SerList, Firm)
                    FirstPart = True

                    txtTitle.Text = ("   Setting Up Job")
                    System.Threading.Thread.Sleep(10)
                    jobSpec = Me.cmbJob.Text
                    DontclickTheButton = True
                    Job = cmbJob.Text
                    Me.Refresh()
                    Me.cmbPart.Text = GetPartNumber()
                    Part = Me.cmbPart.Text
                    PartSpec = Me.cmbPart.Text

                    Me.Refresh()
                    RunningOffsets = False
                    System.Threading.Thread.Sleep(10)
                    JobClicked()
                    Me.Refresh()
                    'RecallCal(1)

                    LoadSpecs()
                    SetupVNA(False, 1)
                    VNAStr = cmbVNA.Text
                    Me.Refresh()
                    'PartClicked()
                    DontclickTheButton = False
                    Me.ckROBOT.Enabled = True
                    SwitchPorts = SQL.GetSpecification("SwitchPorts")
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub cmbJob_Click(sender As Object, e As EventArgs) Handles cmbJob.Click
        Try
            'If Me.cmbJob.Text <> " " And Len(Me.cmbJob.Text) = 9 And Uploading = False Then
            'txtTitle.Text = ("   Setting Up Job")
            'System.Threading.Thread.Sleep(10)
            'DontclickTheButton = True
            'Me.cmbJob.Text = Trim(Me.cmbJob.Text)
            'Job = cmbJob.Text
            'Me.cmbPart.Text = GetPartNumber()
            'Part = Me.cmbPart.Text
            'RunningOffsets = False
            'System.Threading.Thread.Sleep(10)
            'JobClicked()
            'RecallCal(1)
            'SetupVNA(False, 1)
            'LoadSpecs()
            'PartClicked()
            'DontclickTheButton = False
            'End If
        Catch ex As Exception

        End Try

    End Sub
  
    Private Sub JobClicked()
        Dim SQLstr As String
        Dim SwPos As String
        Dim GoodJob As Boolean
        Try
            
            If Me.cmbJob.Text = "Add New" Then
                StatusLog.Items.Add("Opening Specs:" & "" & DateTime.Now.ToString)
                Dim SPEC As New frmSpecifications
                SPEC.StartPosition = FormStartPosition.Manual
                ' SPEC.Location = New Point(Globals.XLocation + NewXLocation, Globals.YLocation + Globals.XSize)
                SPEC.ShowDialog()
                Me.Hide()
                frmSpecifications.ShowDialog()
            End If
            Dim OP As New AddArtWorkRevision
            OP.StartPosition = FormStartPosition.Manual
            OP.Location = New Point(globals.XLocation, globals.YLocation)
            OP.ShowDialog()
            txtArtwork.Text = ArtworkRevision
            Me.Refresh()

            TEST1PASS = True
            TEST2PASS = True
            TEST3PASS = True
            TEST4PASS = True
            TEST5PASS = True
            Test1Fail_bypass = False
            Test1Fail_bypass = False
            Test1Fail_bypass = False
            Test1Fail_bypass = False
            GlobalFail_bypass = False

            If SpecAB_TF Then
                Data4L.Visible = True
                Data4H.Visible = True
                Data4.Visible = False
            Else
                Data4L.Visible = False
                Data4H.Visible = False
                Data4.Visible = True
            End If
            If ISO_TF Then
                Data3L.Visible = True
                Data3H.Visible = True
                Data3.Visible = False
            Else
                Data3L.Visible = False
                Data3H.Visible = False
                Data3.Visible = True
            End If
            If IL_TF Then
                Data1L.Visible = True
                Data1H.Visible = True
                Data1.Visible = False
            Else
                Data1L.Visible = False
                Data1H.Visible = False
                Data1.Visible = True
            End If

            'ILSetDone = False
            SwPos = "          J3(3dB) = SW1, J4(3dB) = SW2, J2(ISO) = SW3"
            SQLstr = "SELECT * from Specifications where JobNumber = '" & Trim(Me.cmbJob.Text) & "'"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    GoodJob = True
                    If Not IsDBNull(dr.Item(1)) Then
                        If dr.Item(1) = "90 DEGREE COUPLER" Or dr.Item(1) = "90 DEGREE COUPLER SMD" Then
                            If dr.Item(1) = "90 DEGREE COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 0
                            SpecType = "90 DEGREE COUPLER"
                            ckTest3.Checked = True
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
                            ckTest3.Visible = True
                            PF3.Visible = True
                            txtOffset3.Visible = True
                            SwPos = "          J3(3dB) = SW1, J4(3dB) = SW2, J2(ISO) = SW3"
                        ElseIf dr.Item(1) = "TRANSFORMER" Or dr.Item(1) = "TRANSFORMER SMD" Then
                            If dr.Item(1) = "TRANSFORMER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            IL_TF = dr.Item(91)
                            SpecIndex = 1
                            SpecType = "TRANSFORMER"
                            SwPos = ""
                            ckTest1.Checked = True
                            ckTest2.Checked = True

                            ckTest3.Checked = False
                            TestLabel3.Visible = False
                            Spec3Min.Visible = False
                            Spec3Max.Visible = False
                            Data3.Visible = False
                            ckTest3.Visible = False
                            PF3.Visible = False
                            txtOffset3.Visible = False
                            ckTest4.Checked = False
                            TestLabel4.Visible = False
                            Spec4Min.Visible = False
                            Spec4Max.Visible = False
                            Data4.Visible = False
                            ckTest4.Visible = False
                            PF4.Visible = False
                            txtOffset4.Visible = False
                            ckTest5.Checked = False
                            TestLabel5.Visible = False
                            Spec5Min.Visible = False
                            Spec5Max.Visible = False
                            Data5.Visible = False
                            ckTest5.Visible = False
                            PF5.Visible = False
                            txtOffset5.Visible = False
                        ElseIf dr.Item(1) = "BALUN" Or dr.Item(1) = "BALUN SMD" Then
                            If dr.Item(1) = "BALUN SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 1
                            SpecType = "BALUN"
                            SwPos = "          J2(6dB) = SW1, J2(6dB) = SW2"
                            ckTest3.Checked = False
                            TestLabel3.Visible = False
                            Spec3Min.Visible = False
                            Spec3Max.Visible = False
                            Data3.Visible = False
                            ckTest3.Visible = False
                            PF3.Visible = False
                            txtOffset3.Visible = False
                        ElseIf dr.Item(1) = "SINGLE DIRECTIONAL COUPLER" Or dr.Item(1) = "SINGLE DIRECTIONAL COUPLER SMD" Then
                            If dr.Item(1) = "SINGLE DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SpecType = "SINGLE DIRECTIONAL COUPLER"
                            ckTest3.Checked = True
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
                            ckTest3.Visible = True
                            PF3.Visible = True
                            txtOffset3.Visible = True
                            SwPos = "          OUT = SW1, CPL = SW2, ISO = SW3"
                        ElseIf dr.Item(1) = "DUAL DIRECTIONAL COUPLER" Or dr.Item(1) = "DUAL DIRECTIONAL COUPLER SMD" Then
                            If dr.Item(1) = "DUAL DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SwPos = "          OUT = SW1, CPL = SW2, RFLD = SW3"
                            SwitchPorts = dr.Item(100)
                            If SwitchPorts = 1 Then SwPos = " OUT = SW1, CPL_J4 = SW2, ISO_J3 = SW3, CPL_J3 = SW4, ISO_J4 = SW5"
                            SpecType = "DUAL DIRECTIONAL COUPLER"
                            ckTest3.Checked = True
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
                            ckTest3.Visible = True
                            PF3.Visible = True
                            txtOffset3.Visible = True
                        ElseIf dr.Item(1) = "BI DIRECTIONAL COUPLER" Or dr.Item(1) = "BI DIRECTIONAL COUPLER SMD" Then
                            If dr.Item(1) = "BI DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SpecType = "BI DIRECTIONAL COUPLER"
                            ckTest3.Checked = True
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
                            ckTest3.Visible = True
                            PF3.Visible = True
                            txtOffset3.Visible = True
                            SwPos = "          OUT = SW1, CPL = SW2, RFLD = SW3"
                            If SpecType = "DUAL DIRECTIONAL COUPLER" And SwitchPorts = 1 Then SwPos = " OUT = SW1, CPL_J4 = SW2, ISO_J3 = SW3, CPL_J3 = SW4, ISO_J4 = SW5"
                        ElseIf dr.Item(1) = "COMBINER/DIVIDER" Or dr.Item(1) = "COMBINER/DIVIDER SMD" Then
                            If dr.Item(1) = "COMBINER/DIVIDER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 3
                            ckTest3.Checked = True
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
                            ckTest3.Visible = True
                            PF3.Visible = True
                            txtOffset3.Visible = True
                            SpecType.Contains("COMBINER/DIVIDER")
                            If dr.Item(9) = 2 Then SwPos = "          J2 = SW1, J3 = SW2"
                            If dr.Item(9) = 3 Then SwPos = "          J2 = SW1, J3 = SW2, J3 = SW4"
                            If dr.Item(9) = 4 Then SwPos = "          J2 = SW1, J3 = SW2, J3 = SW4, J4 SW5"
                            'If dr.Item(13) = 0 Then ckTest3.Checked = False
                        End If
                    Else
                        SpecIndex = 0
                        SpecType = "90 DEGREE COUPLER"
                        ckTest3.Checked = True
                        TestLabel3.Visible = True
                        Spec3Min.Visible = True
                        Spec3Max.Visible = True
                        Data3.Visible = True
                        ckTest3.Visible = True
                        PF3.Visible = True
                        txtOffset3.Visible = True
                        SwPos = "        J3(3dB) = SW1, J4(3dB) = SW2, J2(ISO) = SW3"
                        SMD = False
                    End If
                End While
                ats.Close()
            Else
                'System.Threading.Thread.Sleep(100)
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    GoodJob = True
                    If Not IsDBNull(GetTitle()) Then
                        If drLocal.Item(1) = "90 DEGREE COUPLER" Or drLocal.Item(1) = "90 DEGREE COUPLER SMD" Then
                            If drLocal.Item(1) = "90 DEGREE COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 0
                            SpecType = "90 DEGREE COUPLER"
                            ckTest3.Checked = True
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
                            ckTest3.Visible = True
                            PF3.Visible = True
                            txtOffset3.Visible = True
                            SwPos = "          J3(3dB) = SW1, J4(3dB) = SW2, J2(ISO) = SW3"
                        ElseIf drLocal.Item(1) = "BALUN" Or drLocal.Item(1) = "BALUN SMD" Then
                            If drLocal.Item(1) = "BALUN SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 1
                            SpecType = "BALUN"
                            SwPos = "          J2(6dB) = SW1, J2(6dB) = SW2"
                            ckTest3.Checked = False
                            TestLabel3.Visible = False
                            Spec3Min.Visible = False
                            Spec3Max.Visible = False
                            Data3.Visible = False
                            ckTest3.Visible = False
                            PF3.Visible = False
                            txtOffset3.Visible = False
                            txtOffset3.Enabled = False
                        ElseIf drLocal.Item(1) = "SINGLE DIRECTIONAL COUPLER" Or drLocal.Item(1) = "SINGLE DIRECTIONAL COUPLER SMD" Then
                            If drLocal.Item(1) = "SINGLE DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SpecType = "SINGLE DIRECTIONAL COUPLER"
                            ckTest3.Checked = True
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
                            ckTest3.Visible = True
                            PF3.Visible = True
                            txtOffset3.Visible = True
                            SwPos = "          OUT = SW3, CPL = SW1"
                        ElseIf drLocal.Item(1) = "DUAL DIRECTIONAL COUPLER" Or drLocal.Item(1) = "DUAL DIRECTIONAL COUPLER SMD" Then
                            If drLocal.Item(1) = "DUAL DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SwPos = "          OUT = SW3, CPL = SW1, RFLD = SW2"
                            SpecType = "DUAL DIRECTIONAL COUPLER"
                            If SwitchPorts = 1 Then SwPos = " OUT = SW1, CPL_J4 = SW2, ISO_J3 = SW3, CPL_J3 = SW4, ISO_J4 = SW5"
                            ckTest3.Checked = True
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
                            ckTest3.Visible = True
                            PF3.Visible = True
                            txtOffset3.Visible = True
                        ElseIf drLocal.Item(1) = "BI DIRECTIONAL COUPLER" Or drLocal.Item(1) = "BI DIRECTIONAL COUPLER SMD" Then
                            If drLocal.Item(1) = "BI DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SpecType = "BI DIRECTIONAL COUPLER"
                            ckTest3.Checked = True
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
                            ckTest3.Visible = True
                            PF3.Visible = True
                            txtOffset3.Visible = True
                            SwPos = "          OUT = SW3, CPL = SW1, RFLD = SW2"
                        ElseIf drLocal.Item(1) = "COMBINER/DIVIDER" Or drLocal.Item(1) = "COMBINER/DIVIDER SMD" Then
                            If drLocal.Item(1) = "COMBINER/DIVIDER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 3
                            SpecType.Contains("COMBINER/DIVIDER")
                            If drLocal.Item(9) = 2 Then SwPos = "          J2 = SW1, J3 = SW2"
                            If drLocal.Item(9) = 3 Then SwPos = "          J2 = SW1, J3 = SW2, J3 = SW4"
                            If drLocal.Item(9) = 4 Then SwPos = "          J2 = SW1, J3 = SW2, J3 = SW4, J4 SW5"
                            If drLocal.Item(13) = 0 Then ckTest3.Checked = False
                        End If
                    Else
                        SpecIndex = 0
                        SpecType = "90 DEGREE COUPLER"
                        SwPos = "        J3(3dB) = SW1, J4(3dB) = SW2, J2(ISO) = SW3"
                        SMD = False
                    End If
                End While
                atsLocal.Close()
            End If

            Me.txtTitle.Text = GetTitle() & SwPos
            Me.cmbVNA.Text = ScanGPIB.GetModel
            If Me.cmbVNA.Text = "HP_8753C" Or Me.cmbVNA.Text = "HP_8753E" Then
                If SpecStopFreq > 6000 Then
                    MsgBox(SpecStopFreq & "MHz  Exceeds the Frequency Range of the " & Me.cmbVNA.Text & ".  Please move to capible workstation or choose another Job")
                    Exit Sub
                End If
            End If

            If Me.cmbVNA.Text = "HP_8753C" Then
                If SpecStopFreq > GetVNAFreq() Then
                    MsgBox(SpecStopFreq & "MHz  Exceeds the Calibrated Frequency Range of the " & Me.cmbVNA.Text & ".  Please Calibrate the 6GHZ Range")
                    Tests.CalibrateVNA()
                    Exit Sub
                End If
            End If

            If Me.cmbPart.Text Is Nothing Then Me.cmbPart.Text = GetPartNumber()
            If SpecStartFreq = Nothing Then
                Me.txtStartFreq.Text = GetSpecification("StartFreqMHz")
                SpecStartFreq = Me.txtStartFreq.Text
            Else
                Me.txtStartFreq.Text = SpecStartFreq
            End If
            If SpecStopFreq = Nothing Then
                Me.txtStopFreq.Text = GetSpecification("StopFreqMHz")
                SpecStopFreq = Me.txtStopFreq.Text
            Else
                Me.txtStopFreq.Text = SpecStopFreq
            End If
           

            If IL_TF Then
                Me.Spec1Max.Text = Format(GetSpecification("InsertionLoss"), "0.00") & "/" & Format(GetSpecification("IL_ex"), "0.00")
                SpecILH = Format(GetSpecification("InsertionLoss"), "0.00")
                Me.Spec1Min.Text = "N/A"
                'SpecILL = Me.Spec1Min.Text
                Me.Data1H.Text = ""
                Me.Data1L.Text = ""

            Else
                Me.Spec1Max.Text = Format(GetSpecification("InsertionLoss"), "0.00")
                SpecIL = Me.Spec1Max.Text
                Me.Spec1Min.Text = "N/A"
                Me.Data1.Text = ""
            End If

            Me.Spec2Min.Text = "N/A"
            Me.Data2.Text = ""

            If Not IsDBNull(GetSpecification("VSWR")) Then
                Me.Spec2Max.Text = Format(GetSpecification("VSWR"), "0.0")
                SpecTest2 = GetSpecification("VSWR")
            End If

            If SpecIndex = 0 Or SpecIndex = 1 Or SpecIndex = 3 Then
                If Not IsDBNull(GetSpecification("Isolation")) And ISO_TF Then
                    TestLabel3.Text = "Isolation D/B:  dB"
                    SpecISOL = Format(GetSpecification("IsolationL"), "0.0")
                    SpecISOH = Format(GetSpecification("IsolationH"), "0.0")
                    Me.Spec3Max.Text = 0 - SpecISOL & "/" & 0 - SpecISOH
                    Me.Data3H.Text = ""
                    Me.Data3L.Text = ""
                Else
                    TestLabel3.Text = "Isolation:  dB"
                    Me.Spec3Max.Text = Format(0 - GetSpecification("Isolation"), "0.0")
                    SpecTest3 = Format(GetSpecification("Isolation"), "0.0")
                    SpecISO = SpecTest3
                    Me.Data3.Text = ""
                End If
                Me.Spec3Min.Text = "N/A"
                Me.Spec3Min.ForeColor = Color.LightGray
                Me.Spec3Max.ForeColor = Color.CornflowerBlue

                TestLabel4.Text = "Amplitude Balance dB"
                If Not IsDBNull(GetSpecification("AmplitudeBalance")) Then
                    Me.Spec4Min.Text = Format(GetSpecification("AmplitudeBalance"), "0.00")
                    Me.Spec4Min.ForeColor = Color.CornflowerBlue
                    Me.Spec4Max.ForeColor = Color.CornflowerBlue
                    If SpecAB_TF Then
                        TestLabel4.Text = "Amplitude Balance D/B  dB"
                        SpecTest4 = Me.Spec4Min.Text
                        Me.Spec4Min.Text = Me.Spec4Min.Text & "/" & SpecAB_exp
                        Me.Spec4Max.Text = Me.Spec4Min.Text
                        SpecTest4_exp = SpecAB_exp
                    Else
                        SpecTest4 = Me.Spec4Min.Text
                        Me.Spec4Max.Text = Me.Spec4Min.Text
                    End If
                End If
                If SpecAB_TF Then
                    Data4L.Text = ""
                    Data4H.Text = ""
                Else
                    Data4.Text = ""
                End If
                TestLabel5.Text = "Phase Balance: Deg"
                Dim Temp As String = Format(GetSpecification("PhaseBalance"), "0.0")

                If Not IsDBNull(GetSpecification("PhaseBalance")) Then Me.Spec5Min.Text = Format(GetSpecification("PhaseBalance"), "0.0")
                Me.Data5.Text = ""
                If Not IsDBNull(GetSpecification("PhaseBalance")) Then
                    Me.Spec5Max.Text = Format(GetSpecification("PhaseBalance"), "0.0")
                    Me.Spec5Max.ForeColor = Color.CornflowerBlue
                    Me.Spec5Min.Text = Me.Spec5Max.Text
                    Me.Spec5Min.ForeColor = Color.CornflowerBlue
                    SpecTest5 = Me.Spec5Min.Text
                End If

            ElseIf SpecIndex = 2 Then
                TestLabel3.Text = "Coupling:  dB"
                If Not IsDBNull(GetSpecification("Coupling")) Then
                    Me.Spec3Min.Text = Format(GetSpecification("Coupling") - GetSpecification("CoupPlusMinus"), "0.0")
                    Me.Spec3Min.ForeColor = Color.CornflowerBlue
                    Me.Spec3Max.Text = Format(GetSpecification("Coupling") + GetSpecification("CoupPlusMinus"), "0.0")
                    Me.Spec3Max.ForeColor = Color.CornflowerBlue
                    SpecCOUP = GetSpecification("Coupling")
                End If
                Me.Data3.Text = ""
                TestLabel4.Text = "Directivity: dB"
                If Not IsDBNull(GetSpecification("Directivity")) Then
                    Me.Spec4Min.Text = Format(GetSpecification("Directivity"), "0.0")
                    Me.Spec3Min.ForeColor = Color.CornflowerBlue
                    SpecDIRECT = Me.Spec4Min.Text
                End If
                If SpecAB_TF Then
                    Data4L.Text = ""
                    Data4H.Text = ""
                Else
                    Data4.Text = ""
                End If
                Me.Spec4Max.Text = "N/A"
                Me.Spec4Max.ForeColor = Color.LightGray

                TestLabel5.Text = "Coupled Flatness dB"
                Me.Data5.Text = ""
                If Not IsDBNull(GetSpecification("CoupledFlatness")) Then
                    Me.Spec5Max.Text = Format(GetSpecification("CoupledFlatness"), "0.00")
                    Me.Spec5Min.ForeColor = Color.CornflowerBlue
                    Me.Spec5Min.Text = Format(GetSpecification("CoupledFlatness"), "0.00")
                    Me.Spec5Max.ForeColor = Color.CornflowerBlue
                    SpecCOUPFLAT = Me.Spec5Min.Text
                End If
            End If

            If Not IsDBNull(GetSpecification("Offset1")) Then Me.txtOffset1.Text = GetSpecification("Offset1")
            If Not IsDBNull(GetSpecification("Offset2")) Then Me.txtOffset2.Text = GetSpecification("Offset2")
            If Not IsDBNull(GetSpecification("Offset3")) Then Me.txtOffset3.Text = GetSpecification("Offset3")
            If Not IsDBNull(GetSpecification("Offset4")) Then Me.txtOffset4.Text = GetSpecification("Offset4")
            If Not IsDBNull(GetSpecification("Offset5")) Then Me.txtOffset5.Text = GetSpecification("Offset5")

            If Not IsDBNull(GetSpecification("Test1")) Then Me.ckTest1.Checked = GetSpecification("Test1")
            If Not IsDBNull(GetSpecification("Test2")) Then Me.ckTest2.Checked = GetSpecification("Test2")
            If Not IsDBNull(GetSpecification("Test3")) Then Me.ckTest3.Checked = GetSpecification("Test3")
            If Not IsDBNull(GetSpecification("Test4")) Then Me.ckTest4.Checked = GetSpecification("Test4")
            If Not IsDBNull(GetSpecification("Test5")) Then Me.ckTest5.Checked = GetSpecification("Test5")

            UUTMessage.Text = "  UUT TESTS  --  Load Unit #1"


            resumeTest = ResumeTesting()
            If Job = Nothing Then
                GoodJob = False
                Me.cmbPart.Text = "Part Number"
                Me.txtTitle.Text = "Title"
                Me.txtStartFreq.Text = "Start Freq"
                Me.txtStopFreq.Text = "Stop Freq"

                Me.Spec1Min.Text = "N/A"
                Me.Data1.Text = ""
                Me.Spec1Max.Text = "IL Max"

                Me.Spec2Min.Text = "N/A"
                Me.Data2.Text = ""
                Me.Spec2Max.Text = "RL Max"

                Me.Spec3Min.Text = "TEST3 Max"
                Me.Data3.Text = ""
                Me.Spec3Max.Text = "N/A"

                Me.Spec4Min.Text = "TEST4 Max"
                If SpecAB_TF Then
                    Data4L.Text = ""
                    Data4H.Text = ""
                Else
                    Data4.Text = ""
                End If
                Me.Spec4Max.Text = "TEST4 Min"

                Me.Spec5Min.Text = "TEST5 Max"
                Me.Data5.Text = ""
                Me.Spec5Max.Text = "TEST5 Min"

                UUTMessage.Text = "  UUT TESTS  --  Load Part Number"

            End If

            ResetTests(resumeTest)
            If Not resumeTest Then
                UUTNum = 0
                UUTStatusColor.BackColor = Color.LawnGreen
            End If

            UUTCount.Text = Str(UUTNum)
            If Simulation.Checked = False And GoodJob Then
                RecallCal(1)
                Tests.SetupVNA(True, 1)
            End If

            Exit Sub

        Catch ex As Exception

        End Try


    End Sub
    Private Sub cmbPart_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPart.SelectedIndexChanged
        If Not DontclickTheButton Then
            Me.cmbPart.Text = Trim(Me.cmbPart.GetItemText(Me.cmbPart.SelectedItem))
            Part = cmbPart.Text
            PartClicked()
        End If
    End Sub
    Private Sub cmbPart_Click(sender As Object, e As EventArgs) Handles cmbPart.Click
        Me.cmbPart.Text = Trim(Me.cmbPart.GetItemText(Me.cmbPart.SelectedItem))
        Part = cmbPart.Text
        'PartClicked()
        Me.ckROBOT.Enabled = True
    End Sub

    Public Sub PartClicked()
        Dim SQLstr As String
        Dim SwPos As String
        Dim GoodJob As Boolean
        Try

            If Me.cmbPart.Text = "IPP - " Or Me.cmbPart.Text = " " Then
                Me.cmbPart.Text = " "
                Exit Sub
            End If
            TEST1PASS = True
            TEST2PASS = True
            TEST3PASS = True
            TEST4PASS = True
            TEST5PASS = True

            RunningOffsets = False
            DontclickTheButton = True
            ILSetDone = False
            LoadSpecs()
            SQLstr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "'"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    GoodJob = True
                    If Not IsDBNull(GetPartNumber()) Then
                        If dr.Item(1) = "90 DEGREE COUPLER" Or dr.Item(1) = "90 DEGREE COUPLER SMD" Then
                            If dr.Item(1) = "90 DEGREE COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 0
                            SpecType = "90 DEGREE COUPLER"
                            SwPos = "          J3(3dB) = SW1, J4(3dB) = SW2, J2(ISO) = SW3"
                        ElseIf dr.Item(1) = "BALUN" Or dr.Item(1) = "BALUN SMD" Then
                            If dr.Item(1) = "BALUN SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 1
                            SpecType = "BALUN"
                            SwPos = "          J2(6dB) = SW1, J2(6dB) = SW2"
                            ckTest3.Checked = False
                            TestLabel3.Enabled = False
                            Spec3Min.Enabled = False
                            Spec3Max.Enabled = False
                            Data3.Enabled = False
                            ckTest3.Enabled = False
                            PF3.Enabled = False
                            txtOffset3.Enabled = False
                        ElseIf dr.Item(1) = "SINGLE DIRECTIONAL COUPLER" Or dr.Item(1) = "SINGLE DIRECTIONAL COUPLER SMD" Then
                            If dr.Item(1) = "SINGLE DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SpecType = "SINGLE DIRECTIONAL COUPLER"
                            SwPos = "          OUT = SW3, CPL = SW1"
                        ElseIf dr.Item(1) = "DUAL DIRECTIONAL COUPLER" Or dr.Item(1) = "DUAL DIRECTIONAL COUPLER SMD" Then
                            If dr.Item(1) = "DUAL DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SwPos = "          OUT = SW3, CPL = SW1, RFLD = SW2"
                            SpecType = "DUAL DIRECTIONAL COUPLER"
                            If SwitchPorts = 1 Then SwPos = " OUT = SW1, CPL_J4 = SW2, ISO_J3 = SW3, CPL_J3 = SW4, ISO_J4 = SW5"
                        ElseIf dr.Item(1) = "BI DIRECTIONAL COUPLER" Or dr.Item(1) = "BI DIRECTIONAL COUPLER SMD" Then
                            If dr.Item(1) = "BI DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SpecType = "BI DIRECTIONAL COUPLER"
                            SwPos = "          OUT = SW3, CPL = SW1, RFLD = SW2"
                        ElseIf dr.Item(1) = "COMBINER/DIVIDER" Or dr.Item(1) = "COMBINER/DIVIDER SMD" Then
                            If dr.Item(1) = "COMBINER/DIVIDER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 3
                            SpecType.Contains("COMBINER/DIVIDER")
                            If dr.Item(9) = 2 Then SwPos = "          J2 = SW1, J3 = SW2"
                            If dr.Item(9) = 3 Then SwPos = "          J2 = SW1, J3 = SW2, J3 = SW4"
                            If dr.Item(9) = 4 Then SwPos = "          J2 = SW1, J3 = SW2, J3 = SW4, J4 SW5"
                            If dr.Item(13) = 0 Then ckTest3.Checked = False
                        End If
                    Else
                        SpecIndex = 0
                        SpecType = "90 DEGREE COUPLER"
                        SwPos = "        J3(3dB) = SW1, J4(3dB) = SW2, J2(ISO) = SW3"
                        SMD = False
                    End If
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
                    GoodJob = True
                    If Not IsDBNull(GetSpecification("JobNumber")) Then
                        If drLocal.Item(1) = "90 DEGREE COUPLER" Or drLocal.Item(1) = "90 DEGREE COUPLER SMD" Then
                            If drLocal.Item(1) = "90 DEGREE COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 0
                            SpecType = "90 DEGREE COUPLER"
                            SwPos = "          J3(3dB) = SW1, J4(3dB) = SW2, J2(ISO) = SW3"
                        ElseIf drLocal.Item(1) = "BALUN" Or drLocal.Item(1) = "BALUN SMD" Then
                            If drLocal.Item(1) = "BALUN SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 1
                            SpecType = "BALUN"
                            SwPos = "          J2(6dB) = SW1, J2(6dB) = SW2"
                            ckTest3.Checked = False
                            TestLabel3.Enabled = False
                            Spec3Min.Enabled = False
                            Spec3Max.Enabled = False
                            Data3.Enabled = False
                            ckTest3.Enabled = False
                            PF3.Enabled = False
                            txtOffset3.Enabled = False
                        ElseIf drLocal.Item(1) = "SINGLE DIRECTIONAL COUPLER" Or drLocal.Item(1) = "SINGLE DIRECTIONAL COUPLER SMD" Then
                            If drLocal.Item(1) = "SINGLE DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SpecType = "SINGLE DIRECTIONAL COUPLER"
                            SwPos = "          OUT = SW3, CPL = SW1"
                        ElseIf drLocal.Item(1) = "DUAL DIRECTIONAL COUPLER" Or drLocal.Item(1) = "DUAL DIRECTIONAL COUPLER SMD" Then
                            If drLocal.Item(1) = "DUAL DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SwPos = "          OUT = SW3, CPL = SW1, RFLD = SW2"
                            SpecType = "DUAL DIRECTIONAL COUPLER"
                            If SwitchPorts = 1 Then SwPos = " OUT = SW1, CPL_J4 = SW2, ISO_J3 = SW3, CPL_J3 = SW4, ISO_J4 = SW5"
                        ElseIf drLocal.Item(1) = "BI DIRECTIONAL COUPLER" Or drLocal.Item(1) = "BI DIRECTIONAL COUPLER SMD" Then
                            If drLocal.Item(1) = "BI DIRECTIONAL COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 2
                            SpecType = "BI DIRECTIONAL COUPLER"
                            SwPos = "          OUT = SW3, CPL = SW1, RFLD = SW2"
                        ElseIf drLocal.Item(1) = "COMBINER/DIVIDER" Or drLocal.Item(1) = "COMBINER/DIVIDER SMD" Then
                            If drLocal.Item(1) = "COMBINER/DIVIDER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 3
                            SpecType.Contains("COMBINER/DIVIDER")
                            If drLocal.Item(9) = 2 Then SwPos = "          J2 = SW1, J3 = SW2"
                            If drLocal.Item(9) = 3 Then SwPos = "          J2 = SW1, J3 = SW2, J3 = SW4"
                            If drLocal.Item(9) = 4 Then SwPos = "          J2 = SW1, J3 = SW2, J3 = SW4, J4 SW5"
                            If drLocal.Item(13) = 0 Then ckTest3.Checked = False
                        End If
                    Else
                        SpecIndex = 0
                        SpecType = "90 DEGREE COUPLER"
                        SwPos = "        J3(3dB) = SW1, J4(3dB) = SW2, J2(ISO) = SW3"
                        SMD = False
                    End If
                End While
            End If

            'MSChart.ResetChartData(SpecType)
            'MSChart.UpDateChart(SpecType)

            Me.cmbVNA.Text = ScanGPIB.GetModel
            If Me.cmbVNA.Text = "HP_8753C" Or Me.cmbVNA.Text = "HP_8753E" Then
                If SpecStopFreq > 6000 Then
                    MsgBox(SpecStopFreq & "MHz  Exceeds the Frequency Range of the " & Me.cmbVNA.Text & ".  Please move to capible workstation or choose another Job")
                    Exit Sub
                End If
            End If

            If Me.cmbVNA.Text = "HP_8753C" Then
                If SpecStopFreq > GetVNAFreq() Then
                    MsgBox(SpecStopFreq & "MHz  Exceeds the Calibrated Frequency Range of the " & Me.cmbVNA.Text & ".  Please Calibrate the 6GHZ Range")
                    Tests.CalibrateVNA()
                    Exit Sub
                End If
            End If

            If Me.cmbPart.Text IsNot Nothing Then Me.cmbPart.Text = GetPartNumber()
            If Not IsDBNull(GetSpecification("StartFreqMHz")) Then

                SpecStartFreq = GetSpecification("StartFreqMHz")
            End If
            If SpecStopFreq = Nothing Then SpecStopFreq = GetSpecification("StopFreqMHz")
            Me.txtStopFreq.Text = SpecStopFreq
            Me.Spec1Min.Text = "N/A"
            Me.Data1.Text = ""

            If SpecIL = Nothing Then Me.Spec1Max.Text = GetSpecification("InsertionLoss")
            SpecIL = GetSpecification("InsertionLoss")
            Me.Spec2Min.Text = "N/A"
            Me.Data2.Text = ""
            If Not IsDBNull(GetSpecification("VSWR")) Then
                Me.Spec2Max.Text = Math.Round(VSWRtoRL(CDbl(GetSpecification("VSWR"))), 1)
                SpecTest2 = GetSpecification("VSWR")
            End If

            If Not IsDBNull(GetSpecification("Isolation")) And ISO_TF Then
                TestLabel3.Text = "Isolation D/B:  dB"
                SpecISOL = Format(GetSpecification("IsolationL"), "0.0")
                SpecISOH = Format(GetSpecification("IsolationH"), "0.0")
                Me.Spec3Max.Text = 0 - SpecISOL & "/" & 0 - SpecISOH
                Me.Data3H.Text = ""
                Me.Data3L.Text = ""
            Else
                TestLabel3.Text = "Isolation:  dB"
                Me.Spec3Max.Text = Format(0 - GetSpecification("Isolation"), "0.0")
                SpecTest3 = Format(GetSpecification("Isolation"), "0.0")
                SpecISO = SpecTest3
                Me.Data3.Text = ""
            End If
            Me.Spec3Min.Text = "N/A"
            Me.Spec3Min.ForeColor = Color.LightGray
            Me.Spec3Max.ForeColor = Color.CornflowerBlue

            TestLabel4.Text = "Amplitude Balance dB"
            AB = GetSpecification("AmplitudeBalance")
            Me.Spec4Min.Text = AB
            Me.Spec4Min.ForeColor = Color.CornflowerBlue
            If SpecAB_TF Then
                Me.Spec4Max.Text = AB & "/" & SpecAB_exp
                Me.Spec4Min.Text = AB & "/" & SpecAB_exp
            Else
                Me.Spec4Max.Text = AB
            End If
            Me.Spec4Max.Text = AB
            Me.Spec4Max.ForeColor = Color.CornflowerBlue
            SpecTest4 = AB

            If SpecAB_TF Then
                Data4L.Text = ""
                Data4H.Text = ""
            Else
                Data4.Text = ""
            End If
            TestLabel5.Text = "Phase Balance: Deg"

            Me.Spec5Min.Text = GetSpecification("PhaseBalance")
            Me.Data5.Text = ""
            Me.Spec5Max.Text = PB
            Me.Spec5Max.ForeColor = Color.CornflowerBlue
            Me.Spec5Min.Text = PB
            Me.Spec5Min.ForeColor = Color.CornflowerBlue
            SpecTest5 = PB


            If SpecIndex = 2 Then
                TestLabel3.Text = "Coupling:  dB"
                If Not IsDBNull(GetSpecification("Coupling")) Then
                    Me.Spec3Min.Text = GetSpecification("Coupling") - GetSpecification("CoupPlusMinus")
                    Me.Spec3Min.ForeColor = Color.CornflowerBlue
                    Me.Spec3Max.Text = GetSpecification("Coupling") + GetSpecification("CoupPlusMinus")
                    Me.Spec3Max.ForeColor = Color.CornflowerBlue
                    SpecCOUP = GetSpecification("Coupling")
                End If
                Me.Data3.Text = ""
                TestLabel4.Text = "Directivity: dB"
                If Not IsDBNull(GetSpecification("Directivity")) Then
                    Me.Spec4Min.Text = GetSpecification("Directivity")
                    Me.Spec3Min.ForeColor = Color.CornflowerBlue
                    SpecDIRECT = GetSpecification("Directivity")
                End If
                If SpecAB_TF Then
                    Data4L.Text = ""
                    Data4H.Text = ""
                Else
                    Data4.Text = ""
                End If
                Me.Spec4Max.Text = "N/A"
                Me.Spec4Max.ForeColor = Color.LightGray

                TestLabel5.Text = "Coupled Flatness dB"
                Me.Data5.Text = ""
                SpecCOUPFLAT = GetSpecification("CoupledFlatness")
                Me.Spec5Max.Text = SpecCOUPFLAT
                Me.Spec5Min.ForeColor = Color.CornflowerBlue
                Me.Spec5Min.Text = SpecCOUPFLAT
                Me.Spec5Max.ForeColor = Color.CornflowerBlue

            End If

            Me.txtOffset1.Text = GetSpecification("Offset1")
            Me.txtOffset2.Text = GetSpecification("Offset2")
            Me.txtOffset3.Text = GetSpecification("Offset3")
            Me.txtOffset4.Text = GetSpecification("Offset4")
            Me.txtOffset5.Text = GetSpecification("Offset5")

            Me.ckTest1.Checked = GetSpecification("Test1")
            Me.ckTest2.Checked = GetSpecification("Test2")
            Me.ckTest3.Checked = GetSpecification("Test3")
            Me.ckTest4.Checked = GetSpecification("Test4")
            Me.ckTest5.Checked = GetSpecification("Test5")

            UUTMessage.Text = "  UUT TESTS  --  Load Unit #1"



            If Part = Nothing Then
                GoodJob = False
                Me.cmbPart.Text = "Part Number"
                Me.txtTitle.Text = "Title"
                Me.txtStartFreq.Text = "Start Freq"
                Me.txtStopFreq.Text = "Stop Freq"

                Me.Spec1Min.Text = "N/A"
                Me.Data1.Text = ""
                Me.Spec1Max.Text = "IL Max"

                Me.Spec2Min.Text = "N/A"
                Me.Data2.Text = ""
                Me.Spec2Max.Text = "RL Max"

                Me.Spec3Min.Text = "TEST3 Max"
                Me.Data3.Text = ""
                Me.Spec3Max.Text = "N/A"

                Me.Spec4Min.Text = "TEST4 Max"
                If SpecAB_TF Then
                    Data4L.Text = ""
                    Data4H.Text = ""
                Else
                    Data4.Text = ""
                End If
                Me.Spec4Max.Text = "TEST4 Min"

                Me.Spec5Min.Text = "TEST5 Max"
                Me.Data5.Text = ""
                Me.Spec5Max.Text = "TEST5 Min"

                'ExpectedProgress.Min = 0
                'ExpectedProgress.Max = 150 * 5
                'ExpectedProgress.Value = 0
                UUTMessage.Text = "  UUT TESTS  --  Load Part Number"

            End If

            ResetTests()
            If Not SpecType = Nothing Then
                'MSChart.InitializeChart()
                'MSChart.ResetChartData(SpecType)
                'MSChart.UpDateChart(SpecType)
                UUTNum = 0
                UUTStatusColor.BackColor = Color.LawnGreen
            End If

            UUTCount.Text = Str(UUTNum)
            If Simulation.Checked = False And GoodJob Then Tests.SetupVNA(True, 1)
            Exit Sub

        Catch ex As Exception

        End Try


    End Sub
    Private Sub StarStartExpectedTimeline()
        Dim tempstr(4) As String
        Dim tempstr1(4) As String
        Dim QuantityMinusComplete As Integer

        '********* Expected Time Startup
        ExpectedTimer.Start()
        PauseTimer.Start()

        ExpectedUUTTime = Nothing
        TimeStart = DateTime.Now.ToString
        QuantityMinusComplete = Quantity - UUTCount.Text
        TotalTime = TimeSpan.FromHours(QuantityMinusComplete / PPH)
        TimePerUUT = TimeSpan.FromMinutes(60 / PPH)
        TPP = 60 / PPH
        TPP = Format(Math.Round(TPP, 3), "0.00")




        'Format UUT Time
        UUTTime.Text = TPP.ToString
        'tempstr = Split(UUTTime.Text, ".")
        'tempstr1 = Split(tempstr(0), ":")
        'UUTTime.Text = tempstr1(1) & "." & tempstr1(2)
        '  tempstr1 = Split(tempstr(1), "0")
        ' UUTTime.Text = UUTTime.Text & ":" & tempstr1(0)

        If UUTNum > Quantity Then
            Quantity = UUTNum + Quantity
            ExpectedProgress.Maximum = Quantity
            ActualProgress.Maximum = Quantity
        End If

        ExpectedProgress.Maximum = Quantity
        If UUTNum <= ExpectedProgress.Maximum Then ExpectedProgress.Value = UUTNum



        'Format Expected Time
        ExpectedCompletion.Text = TimeStart + TotalTime - TimePerUUT
        tempstr = Split(ExpectedCompletion.Text, " ")
        ExpectedCompletion.Text = tempstr(1)

        'Format Start Time
        txtStartTime.Text = TimeStart
        tempstr = Split(txtStartTime.Text, " ")
        txtStartTime.Text = tempstr(1)


        If UUTNum <= ExpectedProgress.Maximum Then ExpectedProgress.Value = UUTNum
        ActualProgress.Maximum = Quantity
        If UUTNum <= ActualProgress.Maximum Then ActualProgress.Value = UUTNum

        txtExpected.Text = UUTNum

    End Sub
    Private Sub ExpectedTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExpectedTimer.Tick
        Dim tempstr(4) As String

        'Format Start Time
        txtCurrentTime.Text = DateTime.Now.ToString
        tempstr = Split(txtCurrentTime.Text, " ")
        txtCurrentTime.Text = tempstr(1)

        If ExpectedUUTTime = Nothing Then
            Timernow = DateTime.Now.Ticks - TimeStart.Ticks
            If Timernow >= TimePerUUT.Ticks Then
                ExpectedUUTCount = ExpectedUUTCount + 1
                ExpectedUUTTime = DateTime.Now.Ticks
                If ExpectedProgress.Value + 1 <= ExpectedProgress.Maximum Then ExpectedProgress.Value = ExpectedProgress.Value + 1
                Timernow = DateTime.Now.Ticks - ExpectedUUTTime
                txtExpected.Text = txtExpected.Text + 1
                ExpectedProgress2 = CInt(txtExpected.Text)
            End If
        Else
            Timernow = DateTime.Now.Ticks - ExpectedUUTTime
            If Timernow >= TimePerUUT.Ticks Then
                ExpectedUUTCount = ExpectedUUTCount + 1
                ExpectedUUTTime = DateTime.Now.Ticks
                If ExpectedProgress.Value + 1 <= ExpectedProgress.Maximum Then ExpectedProgress.Value = ExpectedProgress.Value + 1
                txtExpected.Text = txtExpected.Text + 1
                ExpectedProgress2 = CInt(txtExpected.Text)
            End If
        End If
        If ExpectedProgress2 > ActualProgress.Value Then ActualProgress.ForeColor = Color.Red
        If ExpectedProgress2 < ActualProgress.Value Then ActualProgress.ForeColor = Color.LawnGreen
        If ExpectedProgress2 = ActualProgress.Value Then ActualProgress.ForeColor = Color.Yellow
        If ExpectedProgress2 > ExpectedProgress.Maximum Then ExpectedProgress.ForeColor = Color.Red


        If CInt(ExpectedProgress2) > 0 Then
            txtEfficiency.Text = Math.Round(ActualProgress.Value / ExpectedProgress2 * 100, 0) & "%"
        Else
            txtEfficiency.Text = "N/A"
        End If
GetOut:

        If ActualProgress.Value / ExpectedProgress2 * 100 >= 100 Then txtEfficiency.ForeColor = Color.LawnGreen
        If ActualProgress.Value / ExpectedProgress2 * 100 < 100 Then txtEfficiency.ForeColor = Color.Red

        TimeComplete = DateTime.Now - TimeStart



    End Sub
    Private Sub CheckTestResume()
        Dim tempstr(4) As String
        LastTest = Date.Now.Ticks
        If txtCurrentTime.Text = "PAUSED" Then
            StatusLog.Items.Add("Test Restarted:" & "" & DateTime.Now.ToString)
            ExpectedTimer.Start()
            ExpectedCompletion.Text = TimeStart + TotalTime - TimePerUUT + TimeComplete
            tempstr = Split(ExpectedCompletion.Text, " ")
            ExpectedCompletion.Text = tempstr(1)
            If ExpectedProgress.Value - TimePerUUT_Pause >= 0 Then ExpectedProgress.Value = ExpectedProgress.Value - TimePerUUT_Pause + 1
            ExpectedUUTCount = ExpectedUUTCount - TimePerUUT_Pause + 1
            txtExpected.Text = txtExpected.Text - TimePerUUT_Pause + 1
            ExpectedUUTTime = DateTime.Now.Ticks
            ExpectedProgress2 = CInt(txtExpected.Text)
            txtEfficiency.Text = Math.Round(ActualProgress.Value / ExpectedProgress2 * 100, 0) & "%"

            If ExpectedProgress2 > ActualProgress.Value Then ActualProgress.ForeColor = Color.Red
            If ExpectedProgress2 < ActualProgress.Value Then ActualProgress.ForeColor = Color.LawnGreen
            If ExpectedProgress2 = ActualProgress.Value Then ActualProgress.ForeColor = Color.Yellow
            If ExpectedProgress2 > ExpectedProgress.Maximum Then ExpectedProgress.ForeColor = Color.Red

            If ActualProgress.Value / ExpectedProgress2 * 100 >= 100 Then txtEfficiency.ForeColor = Color.LawnGreen
            If ActualProgress.Value / ExpectedProgress2 * 100 < 100 Then txtEfficiency.ForeColor = Color.Red

            If ExpectedProgress2 > 0 Then SQL.UpdateEffeciency("Running", txtEfficiency.Text, Now.Date.ToShortDateString, UUTCount.Text)
        End If
    End Sub
    Private Sub PauseTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PauseTimer.Tick

        Pausenow = Date.Now.Ticks - LastTest
        If Pausenow > (TimePerUUT.Ticks * TimePerUUT_Pause) Then
            ExpectedTimer.Stop()
            If Not txtCurrentTime.Text = "PAUSED" Then
                StatusLog.Items.Add("Test Paused:" & "" & Now.Date.ToShortDateString)
                SQL.UpdateEffeciency("Paused", txtEfficiency.Text, txtCurrentTime.Text, UUTCount.Text)
            End If
            txtCurrentTime.Text = "PAUSED"

        End If
    End Sub
    Private Sub ResetTests(Optional Retest As Boolean = False, Optional resumeTest As Boolean = False)
        Try
            If Not Retest Then UUTFail = 0

            If Not Retest Then
                TEST1FailRetest = 0
                PF1.ForeColor = Color.CornflowerBlue
                PF1.Text = "TBD"
                Data1.Text = ""
            ElseIf Retest And ckTest1.Checked Then
                PF1.ForeColor = Color.CornflowerBlue
                PF1.Text = "TBD"
                Data1.Text = ""
            End If

            If Not Retest Then
                TEST2FailRetest = 0
                PF2.ForeColor = Color.CornflowerBlue
                PF2.Text = "TBD"
                Data2.Text = ""
            ElseIf Retest And ckTest2.Checked Then
                PF2.ForeColor = Color.CornflowerBlue
                PF2.Text = "TBD"
                Data2.Text = ""
            End If

            If Not Retest Then
                TEST3FailRetest = 0
                PF3.ForeColor = Color.CornflowerBlue
                PF3.Text = "TBD"
                Data3.Text = ""
            ElseIf Retest And ckTest3.Checked Then
                PF3.ForeColor = Color.CornflowerBlue
                PF3.Text = "TBD"
                Data3.Text = ""
            End If

            If Not Retest Then
                TEST4FailRetest = 0
                PF4.ForeColor = Color.CornflowerBlue
                PF4.Text = "TBD"
                If SpecAB_TF Then
                    Data4L.Text = ""
                    Data4H.Text = ""
                Else
                    Data4.Text = ""
                End If
            ElseIf Retest And ckTest4.Checked Then
                PF4.ForeColor = Color.CornflowerBlue
                PF4.Text = "TBD"
                If SpecAB_TF Then
                    Data4L.Text = ""
                    Data4H.Text = ""
                Else
                    Data4.Text = ""
                End If
            End If

            If Not Retest Then
                TEST5FailRetest = 0
                PF5.ForeColor = Color.CornflowerBlue
                PF5.Text = "TBD"
                Data5.Text = ""
            ElseIf Retest And ckTest5.Checked Then
                PF5.ForeColor = Color.CornflowerBlue
                PF5.Text = "TBD"
                Data5.Text = ""
            End If

            If Not resumeTest Then
                Me.UUTStatusColor.BackColor = Color.LawnGreen
                Me.StartTestFrame.BackColor = Color.LawnGreen
                Me.ReTestFrame.BackColor = Color.LawnGreen

                Me.UUTLabel.ForeColor = Color.LawnGreen
                Me.UUTCount.ForeColor = Color.LawnGreen
                Me.cmdRetest.Enabled = False
                Me.EraseTest.Enabled = False
            End If
        Catch
        End Try

    End Sub

    Private Sub ResetLot()
        'MSChart.ResetChartData(SpecType)
        'MSChart.UpDateChart(SpecType)
        Me.UUTStatusColor.BackColor = Color.LawnGreen
        Me.StartTestFrame.BackColor = Color.LawnGreen
        Me.ReTestFrame.BackColor = Color.LawnGreen
        UUTNum = 0
        GlobalFail = 0
        Test1Fail = 0
        Test2Fail = 0
        TEST3Fail = 0
        TEST4Fail = 0
        TEST5Fail = 0
        Retest1Fail = 0
        Retest2Fail = 0
        RetEST3Fail = 0
        RetEST4Fail = 0
        RetEST5Fail = 0
        LOTFail = 0
        Me.LotFailure.Text = "0"
        Me.UUTCount.Text = "0"
        Me.UUTCount.ForeColor = Color.CornflowerBlue
        Me.UUTLabel.ForeColor = Color.LawnGreen
        Me.cmdStartTest.Text = "Start Test"
        Me.cmdStartTest.Enabled = True
        Me.EraseTest.Enabled = False

        Me.Failures1.ForeColor = Color.LawnGreen
        Me.Failures1.Text = "0.0 %"
        Me.FailTotal1.ForeColor = Color.LawnGreen
        Me.FailTotal1.Text = "0"
        Me.Total1.Text = "0"
        Me.PF1.ForeColor = Color.CornflowerBlue
        Me.PF1.Text = "TBD"
        Me.Data1.Text = ""

        Me.Failures2.ForeColor = Color.LawnGreen
        Me.Failures2.Text = "0.0 %"
        Me.FailTotal2.ForeColor = Color.LawnGreen
        Me.FailTotal2.Text = "0"
        Me.Total2.Text = "0"
        Me.PF2.ForeColor = Color.CornflowerBlue
        Me.PF2.Text = "TBD"
        Me.Data2.Text = ""
        If Not SpecType = "TRANSFORMER" Then
            Me.Data3.Visible = True
            Me.Spec3Min.Visible = True
            Me.Spec3Max.Visible = True
            Me.TestLabel3.Visible = True
            Me.txtOffset3.Visible = True
            Me.ckTest3.Visible = True
            Me.PF3.Visible = True
            Me.Data4.Visible = True
            Me.Spec4Min.Visible = True
            Me.Spec4Max.Visible = True
            Me.TestLabel4.Visible = True
            Me.txtOffset4.Visible = True
            Me.ckTest4.Visible = True
            Me.PF4.Visible = True
            Me.Data5.Visible = True
            Me.Spec5Min.Visible = True
            Me.Spec5Max.Visible = True
            Me.TestLabel5.Visible = True
            Me.txtOffset5.Visible = True
            Me.ckTest5.Visible = True
            Me.PF5.Visible = True
        End If
        Me.Failures3.ForeColor = Color.LawnGreen
        Me.Failures3.Text = "0.0 %"
        Me.FailTotal3.ForeColor = Color.LawnGreen
        Me.FailTotal3.Text = "0"
        Me.Total3.Text = "0"
        Me.PF3.ForeColor = Color.CornflowerBlue
        Me.PF3.Text = "TBD"
        Me.Data3.Text = ""

        Me.Failures4.ForeColor = Color.LawnGreen
        Me.Failures4.Text = "0.0 %"
        Me.FailTotal4.ForeColor = Color.LawnGreen
        Me.FailTotal4.Text = "0"
        Me.Total4.Text = "0"
        Me.PF4.ForeColor = Color.CornflowerBlue
        Me.PF4.Text = "TBD"
        Me.Data4.Text = ""

        Me.Failures5.ForeColor = Color.LawnGreen
        Me.Failures5.Text = "0.0 %"
        Me.FailTotal5.ForeColor = Color.LawnGreen
        Me.FailTotal5.Text = "0"
        Me.Total5.Text = "0"
        Me.PF5.ForeColor = Color.CornflowerBlue
        Me.PF5.Text = "TBD"
        Me.Data5.Text = ""

        ClearFailureLog()
        ClearStatusLog()
        Me.cmdRetest.Enabled = False
        Me.EraseTest.Enabled = False
        Me.LotTestFrame.BackColor = Color.CornflowerBlue
        Me.LotFailureFrame.BackColor = Color.CornflowerBlue

        ExpectedProgress.Value = 0
        ActualProgress.Value = 0
        EndLot.Enabled = False
        Me.EraseJob.Enabled = False

    End Sub

    Private Sub cmbSwitch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSwitch.SelectedIndexChanged
        Dim status As Integer
        Dim StatusRet As Integer

        SwitchCom.Connect()
        ILSetDone = False

        If Me.cmbSwitch.Text = "Switch POS 1" Then
            status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 1:" & "" & DateTime.Now.ToString)
        ElseIf Me.cmbSwitch.Text = "Switch POS 2" Then
            status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 2:" & "" & DateTime.Now.ToString)
        ElseIf Me.cmbSwitch.Text = "Switch POS 3" Then
            status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 3:" & "" & DateTime.Now.ToString)
        ElseIf Me.cmbSwitch.Text = "Switch POS 4" Then
            status = SwitchCom.SetSwitchPosition(4) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 4:" & "" & DateTime.Now.ToString)
        ElseIf Me.cmbSwitch.Text = "Switch POS 5" Then
            status = SwitchCom.SetSwitchPosition(5) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 4:" & "" & DateTime.Now.ToString)
        ElseIf Me.cmbSwitch.Text = "Switch POS 6" Then
            status = SwitchCom.SetSwitchPosition(6) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 4:" & "" & DateTime.Now.ToString)
        End If
        ClearStatusLog()
    End Sub
    Private Sub cmbVNA_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbVNA.SelectedIndexChanged
        VNAClicked()
    End Sub
    Private Sub cmbVNA_Click(sender As Object, e As EventArgs) Handles cmbVNA.Click
        'VNAClicked()
    End Sub

    Public Sub VNAClicked()
        Dim SQLstr As String

        Try
            If cmbVNA.Text = "" Then Exit Sub
            VNAStr = cmbVNA.Text
            ILSetDone = False
            SQLstr = "select * from WorkStation where ComputerName = '" & GetComputerName() & "'"
            If SQL.CheckforRow(SQLstr, "LocalSpecs") = 0 Then
                SQLstr = "Insert Into WorkStation (ComputerName) values ('" & GetComputerName() & "')"
                SQL.ExecuteSQLCommand(SQLstr, "LocalSpecs")
            End If

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE WorkStation Set VNAType = '" & cmbVNA.Text & "' where ComputerName = '" & GetComputerName() & "'"
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.CommandText = "UPDATE WorkStation Set VNAType = '" & cmbVNA.Text & "' where ComputerName = '" & GetComputerName() & "'"
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If

        Catch ex As Exception

        End Try
    End Sub

    Public Function GetVNAType() As String
        Dim SQLstr As String
        Try
            GetVNAType = ""
            SQLstr = "select * from WorkStation where ComputerName = '" & GetComputerName() & "'"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If Not IsNothing(dr.Item(3)) Then GetVNAType = dr.Item(3)
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
                    If Not IsNothing(drLocal.Item(3)) Then GetVNAType = drLocal.Item(3)
                End While
                atsLocal.Close()
            End If
        Catch ex As Exception
            GetVNAType = ""

        End Try
    End Function

    Private Sub cmdRetest_Click(sender As Object, e As EventArgs) Handles cmdRetest.Click
        Dim PassFail As String = "Pass"
        Dim SQLstr As String
        Dim TestID As Long
        Dim SwPos As String
        Dim Removed As Boolean = False
        Dim Passed As Boolean = True
        Dim RetrnStr As String

        TestRunning = True
        TestComplete = False
        Retest_bypass = False
        If SpecAB_TF Then
            Data4L.Visible = True
            Data4H.Visible = True
            Data4.Visible = False
        Else
            Data4L.Visible = False
            Data4H.Visible = False
            Data4.Visible = True
        End If

        If Not Simulation.Checked And Not Connected Then
            ScanGPIB.connect("GPIB0::16::INSTR", GetTimeout())
            Connected = True
        End If

        SwPos = "          J3(3dB) = SW1, J4(3dB) = SW2, J2(ISO) = SW3"
        If DontclickTheButton = True Then Exit Sub
        If SpecType = "DUAL DIRECTIONAL COUPLER" Or (SpecType = "SINGLE DIRECTIONAL COUPLER" And Me.ckTest2.Checked) Then
            If SpecType = "DUAL DIRECTIONAL COUPLER" Then SwPos = "          J2(OUT) = SW3, J3(CLP) = SW1, J4(RFLD) = SW2"
            If SpecType = "DUAL DIRECTIONAL COUPLER" And SwitchPorts = 1 Then SwPos = " OUT = SW1, CPL_J4 = SW2, ISO_J3 = SW3, CPL_J3 = SW4, ISO_J4 = SW5"
            If SpecType = "SINGLE DIRECTIONAL COUPLER" Then SwPos = "          J2(OUT) = SW3, J3(CLP) = SW1"
            txtTitle.Text = SpecType & SwPos
            'MsgBox("Please turn Directional Coupler in Forward direction")
        End If
        Me.Refresh()

        RetestSelect()

        Me.Refresh()
        'StatusLog.Items.Add("Re-Testing UUT:" & UUTCount.Text & "" & DateTime.Now.ToString)
        If Me.cmdRetest.Text = "Re - Test" Then
            ResetTests(True)
            Me.RunStatus.ForeColor = Color.Red
            If UUTCount.Text <> Str(UUTNum) Then
                LastUUT = UUTNum
                UUTNum = UUTCount.Text
            End If
            ILSetDone = False

            If TweakMode Then
                Me.UUTMessage.Text = "  UUT TESTS  --  Testing Undisclosed Unit "
            ElseIf Not TraceChecked Then
                UUTMessage.Text = "  UUT TESTS Marker Mode  --   Load Unit #" & UUTNum + 1
            Else
                Me.UUTMessage.Text = "  UUT TESTS  --  Testing Unit #" & UUTNum
            End If
            Me.cmdStartTest.Text = "UUT" & UUTNum
            Me.cmdStartTest.Enabled = False
            Job = Me.cmbJob.Text
            SerialNumber = "UUT" & UUTNum_Reset
            SQLstr = "select * from TestData where JobNumber = '" & Me.cmbJob.Text & "' and WorkStation = '" & GetComputerName() & "' And SerialNumber = 'UUT" & UUTNum_Reset & "' and artwork_rev = '" & ArtworkRevision & "'"
            If SQL.CheckforRow(SQLstr, "NetworkData") = 0 Then
                SQLstr = "Insert Into TestData (JobNumber, PartNumber,SerialNumber,WorkStation,artwork_rev) values ('" & Job & "','" & Part & "','" & SerialNumber & "','" & GetComputerName() & "','" & ArtworkRevision & "')"
                SQL.ExecuteSQLCommand(SQLstr, "NetworkData")
            End If
            SQLstr = "select * from TestData where JobNumber = '" & Me.cmbJob.Text & "' and WorkStation = '" & GetComputerName() & "' And SerialNumber = 'UUT" & UUTNum_Reset & "' and artwork_rev = '" & ArtworkRevision & "'"
            TestID = SQL.GetTestID(SQLstr, "NetworkData")
            FailCount = 0

            'Insertion Loss
            If Me.ckTest1.Checked Then
                If Me.ckROBOT.Checked Then RobotStatus()
                If TraceChecked Then
                    If (SpecType = "TRANSFORMER") And IL_TF Then PassFail = Tests.InsertionLossTRANS_multiband(, TestID)
                    If SpecType = "TRANSFORMER" Then PassFail = Tests.InsertionLossTRANS(, TestID)
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN") And SpecAB_TF Then PassFail = Tests.InsertionLoss3dB_multiband(, TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.InsertionLoss3dB(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.InsertionLossCOMB(, TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.InsertionLossDIR(, TestID)
                ElseIf Not MutiCalChecked Then
                    If (SpecType = "TRANSFORMER") And IL_TF Then PassFail = Tests.InsertionLossTRANS_multiband(, TestID)
                    If SpecType = "TRANSFORMER" Then PassFail = Tests.InsertionLossTRANS_Marker(, TestID)
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And SpecAB_TF Then PassFail = Tests.InsertionLoss3dB_multiband(, TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.InsertionLoss3dB_marker(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.InsertionLossCOMB_Marker(, TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.InsertionLossDIR_Marker(, TestID)
                End If


                If IL_TF Then
                    RetrnVal1 = IL1
                    SaveTestData("InsertionLoss", RetrnVal1)
                    RetrnVal = IL2
                    SaveTestData("InsertionLoss2", RetrnVal)
                    Data1L.Text = IL1
                    Data1H.Text = IL2
                Else
                    RetrnVal = IL
                    SaveTestData("InsertionLoss", RetrnVal)
                    RetrnStr = CStr(TruncateDecimal(RetrnVal, 2))
                    Data1.Text = Format(RetrnVal, "0.00")
                End If
                status("Blue", "TEST1")
                PF1.Text = PassFail

                If PassFail = "Pass" Then
                    TEST1PASS = True
                    status("Green", "TEST1", True)
                    PF1.Text = PassFail
                    If Test1Fail > 0 Then Test1Fail = Test1Fail - 1
                    If Retest1Fail > 0 Then Retest1Fail = 0
                    Me.Refresh()
                ElseIf PassFail = "Fail" Then
                    status("Red", "TEST1", True)
                    Retest1Fail = Retest1Fail + 1
                    If Retest1Fail > retestFailMax And Not Test1Fail_bypass And Not Master_bypass And Not Retest_bypass Then
                        Test1Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation + 50)
                        ERR.ShowDialog()
                        Test1Failing = False
                        ErrorTimer.Stop()
                        Retest_bypass = True
                    End If
                    Passed = False
                    TEST1PASS = False
                    PF1.Text = PassFail
                    status("Red", "TEST1", True)
                    UUTFail = 1
                End If
                Me.Total1.Text = UUTNum
            Else
                'TEST2PASS = True
                'status("Blue", "TEST1")
                'PF1.Text = "Pass"
                'Me.Total1.Text = UUTNum
            End If


            'Return Loss
            If Me.ckTest2.Checked Then
                PassFail = Tests.ReturnLoss(, TestID)
                RetrnVal = RL
                SaveTestData("ReturnLoss", RetrnVal)
                'status("Blue", "TEST2")
                Data2.Text = Format(RetrnVal, "0.0")
                If PassFail = "Pass" Then
                    TEST2PASS = True
                    status("Green", "TEST2", True)
                    PF2.Text = PassFail
                    If Test2Fail > 0 Then Test2Fail = Test2Fail - 1
                    If Retest2Fail > 0 Then Retest2Fail = 0
                ElseIf PassFail = "Fail" Then
                    Passed = False
                    status("Red", "TEST2", True)
                    Data4.Text = Format(RetrnVal, "0.00")
                    PF4.Text = PassFail
                    Retest2Fail = Retest2Fail + 1
                    If Retest2Fail > retestFailMax And Not Test2Fail_bypass And Not Master_bypass And Not Retest_bypass Then
                        Test2Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation + 50)
                        ERR.ShowDialog()
                        Test2Failing = False
                        Retest_bypass = True
                        ErrorTimer.Stop()
                    End If
                    TEST2PASS = False
                    PF2.Text = PassFail
                    status("Red", "TEST2", True)
                    UUTFail = 1
                End If
                Me.Total2.Text = UUTNum
            Else
                'TEST2PASS = True
                'status("Blue", "TEST2")
                'PF2.Text = "Pass"
                'Me.Total2.Text = UUTNum
            End If
            If SpecType <> "COMBINER/DIVIDER" Then GoTo Test2Sub
Test2SubRet:

            'AmplitudeBalance
            'Directivity
            If Me.ckTest4.Checked Then
                If TraceChecked Then
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN") And SpecAB_TF Then PassFail = Tests.AmplitudeBalance_multiband(, TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN" Then PassFail = Tests.AmplitudeBalance(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.AmplitudeBalanceCOMB(, TestID)
                    If SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Then PassFail = Tests.Directivity(1, SpecType, , TestID)
                ElseIf Not MutiCalChecked Then
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN") And SpecAB_TF Then PassFail = Tests.AmplitudeBalance_multiband(, TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN" Then PassFail = Tests.AmplitudeBalance_Marker(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.AmplitudeBalanceCOMB_Marker(, TestID)
                    If SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Then PassFail = Tests.Directivity_Marker(1, SpecType, , TestID)
                End If
                If SpecType <> "SINGLE DIRECTIONAL COUPLER" Then
                    If SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Then
                        RetrnVal = DIR
                        SaveTestData("Directivity", RetrnVal)
                    Else
                        If SpecAB_TF Then
                            RetrnVal = AB1
                            SaveTestData("AmplitudeBalance1", RetrnVal)
                            RetrnVal = AB2
                            SaveTestData("AmplitudeBalance2", RetrnVal)
                            'remove later
                            RetrnVal = AB
                            SaveTestData("AmplitudeBalance", RetrnVal)
                        Else
                            RetrnVal = AB
                            SaveTestData("AmplitudeBalance", RetrnVal)
                        End If
                    End If
                    status("Blue", "TEST4")
                    If SpecAB_TF Then
                        AB1 = Format(AB1, "0.00")
                        AB2 = Format(AB2, "0.00")
                        Data4L.Text = AB1
                        Data4H.Text = AB2
                        If PassFail = "Pass" Then
                            TEST4PASS = True
                            status("Green", "TEST4L", True)
                            status("Green", "TEST4H", True)
                            If SpecType <> "DUAL DIRECTIONAL COUPLER" And SpecType <> "SINGLE DIRECTIONAL COUPLER" Then
                                TEST4PASS = True
                            End If
                        ElseIf PassFail = "Fail" Then
                            TEST4PASS = False
                            If AB1Pass = "Pass" Then
                                status("Green", "TEST4L", True)
                            ElseIf AB1Pass = "Fail" Then
                                status("Red", "TEST4L", True)
                            End If
                            If AB2Pass = "Pass" Then
                                status("Green", "TEST4H", True)
                            ElseIf AB2Pass = "Fail" Then
                                status("Red", "TEST4H", True)
                            End If
                            If SpecType <> "DUAL DIRECTIONAL COUPLER" And SpecType <> "SINGLE DIRECTIONAL COUPLER" Then
                                TEST4PASS = False
                                TEST4Fail = TEST4Fail + 1
                                If Not GlobalFailed Then
                                    GlobalFail = GlobalFail + 1
                                    GlobalFailed = True
                                End If
                                If GlobalFail > GlobalFailMax And Not GlobalFail_bypass And Not Master_bypass Then
                                    GlobalFailing = True
                                    ErrorTimer.Start()
                                    Dim ERR As New Failures_Manager
                                    ERR.StartPosition = FormStartPosition.Manual
                                    ERR.ShowDialog()
                                    ErrorTimer.Stop()
                                End If
                                If TEST4Fail > TestFailMax And Not TEST4Fail_bypass And Not Master_bypass Then
                                    TEST4Failing = True
                                    ErrorTimer.Start()
                                    Dim ERR As New Failures_Manager
                                    ERR.StartPosition = FormStartPosition.Manual
                                    ERR.Location = New Point(globals.XLocation, globals.YLocation)
                                    ERR.ShowDialog()
                                    TEST4Failing = False
                                    ErrorTimer.Stop()
                                End If
                                status("Red", "TEST4")
                                TEST4FailRetest = TEST4FailRetest + 1
                                UUTFail = 1
                            End If
                        End If
                    Else
                        If PassFail = "Pass" Then
                            status("Green", "TEST4")
                            Data4.Text = Format(RetrnVal, "0.00")
                            PF4.Text = PassFail
                            TEST4PASS = True
                        ElseIf PassFail = "Fail" Then
                            status("Red", "TEST4")
                            TEST4PASS = False
                            Data4.Text = Format(RetrnVal, "0.00")
                            PF4.Text = PassFail
                            TEST4Fail = TEST4Fail + 1
                            Me.Failures4.Text = FormatPercent(((TEST4Fail / UUTNum)), 1)
                            Me.Total4.Text = UUTNum
                            Me.FailTotal4.Text = TEST4Fail
                            If Not GlobalFailed Then
                                GlobalFail = GlobalFail + 1
                                GlobalFailed = True
                            End If
                            If GlobalFail > GlobalFailMax And Not GlobalFail_bypass And Not Master_bypass Then
                                GlobalFailing = True
                                ErrorTimer.Start()
                                Dim ERR As New Failures_Manager
                                ERR.StartPosition = FormStartPosition.Manual
                                ERR.ShowDialog()
                                ErrorTimer.Stop()
                            End If
                            If TEST4Fail > TestFailMax And Not TEST4Fail_bypass And Not Master_bypass Then
                                TEST4Failing = True
                                ErrorTimer.Start()
                                Dim ERR As New Failures_Manager
                                ERR.StartPosition = FormStartPosition.Manual
                                ERR.Location = New Point(globals.XLocation, globals.YLocation)
                                ERR.ShowDialog()
                                TEST4Failing = False
                                ErrorTimer.Stop()
                            End If
                            status("Red", "TEST4")
                            TEST4FailRetest = TEST4FailRetest + 1
                            UUTFail = 1
                        End If
                    End If
                End If
            End If
            Me.Refresh()

            'PhaseBalance
            'CoupledFlatness
            If Me.ckTest5.Checked Then
                If TraceChecked Then
                    If SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN" Then PassFail = Tests.PhaseBalance(SpecType, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.PhaseBalanceCOMB(SpecType, , TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.CoupledFlatness(1, SpecType, , TestID)
                ElseIf Not MutiCalChecked Then
                    If SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN" Then PassFail = Tests.PhaseBalance_Marker(SpecType, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.PhaseBalanceCOMB_Marker(SpecType, , TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.CoupledFlatness(1, SpecType, , TestID)
                End If


                If SpecType.Contains("DIRECTIONAL") Then
                    RetrnVal = COuP
                Else
                    RetrnVal = PB
                End If
                SaveTestData("PhaseBalance", RetrnVal)
                status("Blue", "TEST5")
                Data5.Text = Format(RetrnVal, "0.00")
                If PassFail = "Pass" Then
                    TEST5PASS = True
                    status("Green", "TEST5", True)
                    PF5.Text = PassFail
                    If Not SpecType.Contains("DIRECTIONAL") Then If TEST5Fail > 0 Then TEST5Fail = TEST5Fail - 1
                    If RetEST5Fail > 0 Then RetEST5Fail = 0
                ElseIf PassFail = "Fail" Then
                    Passed = False
                    status("Red", "TEST5", True)
                    RetEST5Fail = RetEST5Fail + 1
                    If RetEST5Fail > retestFailMax And Not TEST5Fail_bypass And Not Master_bypass And Not Retest_bypass Then
                        TEST5Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation + 50)
                        ERR.ShowDialog()
                        TEST5Failing = False
                        Retest_bypass = True
                        ErrorTimer.Stop()
                    End If
                    TEST5PASS = False
                    PF5.Text = PassFail
                    status("Red", "TEST5", True)
                    UUTFail = 1
                End If
                Me.Total5.Text = UUTNum
            Else
                'TEST5PASS = True
                'status("Blue", "TEST5")
                'PF5.Text = "Pass"
                'Me.Total5.Text = UUTNum
            End If
            If SpecType <> "COMBINER/DIVIDER" Then GoTo TestComplete
Test2Sub:
            'Isolation
            'Coupling
            If Me.ckTest3.Checked Then
                If TraceChecked Then
                    If SpecType = "90 DEGREE COUPLER" Then PassFail = Tests.Isolation(, TestID)
                    If SpecType = "BALUN" Then PassFail = "N/A"
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.IsolationCOMB(, TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.Coupling(1, SpecType, , TestID)
                ElseIf Not MutiCalChecked Then
                    If SpecType = "90 DEGREE COUPLER" Then PassFail = Tests.Isolation_Marker(, TestID)
                    If SpecType = "BALUN" Then PassFail = "Pass"
                    If SpecType.Contains("COMBINER/DIVIDER") And ISO_TF Then PassFail = Tests.IsolationCOMB(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.IsolationCOMB_Marker(, TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.Coupling(1, SpecType, , TestID)
                End If
                If SpecType.Contains("COMBINER/DIVIDER") And ISO_TF Then
                    ISoL = Format(ISoL, "0.00")
                    ISoH = Format(ISoH, "0.00")
                    Data3L.Text = ISoL
                    Data3H.Text = ISoH
                    PF3.Text = PassFail
                    If PassFail = "Pass" Then
                        TEST3PASS = True
                        status("Green", "TEST3L", True)
                        status("Green", "TEST3H", True)
                    ElseIf PassFail = "Fail" Then
                        TEST3PASS = False
                        If AB1Pass = "Pass" Then
                            status("Green", "TEST3L", True)
                        ElseIf AB1Pass = "Fail" Then
                            status("Red", "TEST3L", True)
                        End If
                        If AB2Pass = "Pass" Then
                            status("Green", "TEST3H", True)
                        ElseIf AB2Pass = "Fail" Then
                            status("Red", "TEST3H", True)
                        End If
                    End If

                    If SQLAccess Then
                        SaveTestData("IsolationL", ISoL)
                        RetrnVal = CStr(TruncateDecimal(ISoL, 1))
                        SaveTestData("IsolationH", ISoH)
                        RetrnVal = CStr(TruncateDecimal(ISoH, 1))
                    Else
                        SaveTestData("IsoL", ISoL)
                        RetrnVal = CStr(TruncateDecimal(ISoL, 1))
                        SaveTestData("IsoH", ISoH)
                        RetrnVal = CStr(TruncateDecimal(ISoH, 1))
                    End If
                ElseIf SpecType = "90 DEGREE COUPLER" Then
                    RetrnVal = ISo
                Else
                    RetrnVal = COuP
                End If
                If SpecType = "BALUN" Then
                    'DO NOTHING
                ElseIf SpecType = "BI DIRECTIONAL COUPLER" Then
                    SaveTestData("Coupling", RetrnVal)
                ElseIf SpecType <> "DUAL DIRECTIONAL COUPLER" And SpecType <> "SINGLE DIRECTIONAL COUPLER" Then
                    'DO NOTHING
                Else
                    If SQLAccess Then
                        SaveTestData("Isolation", RetrnVal)
                    Else
                        SaveTestData("Iso", RetrnVal)
                    End If
                End If
                status("Blue", "TEST3")
                Data3.Text = Format(RetrnVal, "0.0")
                If PassFail = "Pass" Then

                ElseIf PassFail = "Pass" Then
                    TEST3PASS = True
                    status("Green", "TEST3", True)
                    PF3.Text = PassFail
                    If Not SpecType.Contains("DIRECTIONAL") Then If TEST3Fail > 0 Then TEST3Fail = TEST3Fail - 1
                    If RetEST3Fail > 0 Then RetEST3Fail = 0
                ElseIf PassFail = "Fail" Then
                    Passed = False
                    status("Red", "TEST3", True)
                    RetEST3Fail = RetEST3Fail + 1
                    If RetEST3Fail > retestFailMax And Not TEST3Fail_bypass And Not Master_bypass And Not Retest_bypass Then
                        TEST3Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation + 50)
                        ERR.ShowDialog()
                        TEST3Failing = False
                        Retest_bypass = True
                        ErrorTimer.Stop()
                    End If
                    TEST3PASS = False
                    PF3.Text = PassFail
                    status("Red", "TEST3", True)
                    UUTFail = 1
                End If
                Me.Total3.Text = UUTNum
            Else
                'TEST3PASS = True
                'status("Blue", "TEST3")
                'PF3.Text = "Pass"
                'Me.Total3.Text = UUTNum
            End If
            If SpecType <> "COMBINER/DIVIDER" Then GoTo Test2SubRet
TestComplete:  ' For everything except Directional Couplers

            'Directonal Couplers reverse direction
            If Not TweakMode And (SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or (SpecType = "SINGLE DIRECTIONAL COUPLER" And Me.ckTest4.Checked)) Then
                If SpecType = "DUAL DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, REFL = SW3"
                If SpecType = "DUAL DIRECTIONAL COUPLER" And SwitchPorts = 1 Then SwPos = " OUT = SW1, CPL_J4 = SW2, ISO_J3 = SW3, CPL_J3 = SW4, ISO_J4 = SW5"
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, ISO = SW3"
                If SpecType = "BI DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, REFL = SW3"
                txtTitle.Text = SpecType & SwPos
                'MsgBox("Please turn the Directional Coupler in the Reverse direction")
            End If

            'Reverse Coupling
            If SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" And Me.ckTest3.Checked Then
                PassFail = Tests.Coupling(2, SpecType, , TestID)
                RetrnVal = COuP
                SaveTestData("Coupling", RetrnVal)
                status("Blue", "TEST3")
                Data3.Text = Format(RetrnVal, "0.0")
                If PassFail = "Pass" Then
                    TEST3PASS = True
                    status("Green", "TEST3", True)
                    PF3.Text = PassFail
                    If TEST3Fail > 0 Then TEST3Fail = TEST3Fail - 1
                    If RetEST3Fail > 0 Then RetEST3Fail = 0
                ElseIf PassFail = "Fail" Then
                    Passed = False
                    status("Red", "TEST4", True)
                    If RetEST3Fail > retestFailMax And Not TEST3Fail_bypass And Not Master_bypass And Not Retest_bypass Then
                        TEST3Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation + 50)
                        ERR.ShowDialog()
                        TEST3Failing = False
                        Retest_bypass = True
                        ErrorTimer.Stop()
                    End If
                    TEST3PASS = False
                    PF3.Text = PassFail
                    status("Red", "TEST3", True)
                    UUTFail = 1
                End If
                Me.Total3.Text = UUTNum
            Else
                'TEST3PASS = True
                'status("Blue", "TEST3")
                'PF3.Text = "Pass"
                'Me.Total3.Text = UUTNum
            End If

            ' Reverse Directivity
            If (SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "SINGLE DIRECTIONAL COUPLER") And Me.ckTest4.Checked Then
                If TraceChecked And Not TweakMode Then
                    PassFail = Tests.Directivity(2, SpecType, , TestID)
                Else
                    PassFail = Tests.Directivity_Marker(2, SpecType, , TestID)
                End If
                RetrnVal = DIR
                SaveTestData("Directivity", RetrnVal)
                status("Blue", "TEST4")
                Data4.Text = Format(RetrnVal, "0.0")
                If PassFail = "Pass" Then
                    TEST4PASS = True
                    status("Green", "TEST4", True)
                    PF4.Text = PassFail
                    If TEST4Fail > 0 Then TEST4Fail = TEST4Fail - 1
                    If RetEST4Fail > 0 Then RetEST4Fail = 0
                ElseIf PassFail = "Fail" Then
                    Passed = False
                    status("Red", "TEST4", True)
                    If RetEST4Fail > retestFailMax And Not TEST4Fail_bypass And Not Master_bypass And Not Retest_bypass Then
                        TEST4Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation + 50)
                        ERR.ShowDialog()
                        TEST4Failing = False
                        Retest_bypass = True
                        ErrorTimer.Stop()
                    End If
                    TEST4PASS = False
                    PF4.Text = PassFail
                    status("Red", "TEST4", True)
                    UUTFail = 1
                End If
                Me.Total4.Text = UUTNum
            Else
                'TEST4PASS = True
                'status("Blue", "TEST4")
                'PF4.Text = "Pass"
                'Me.Total4.Text = UUTNum
            End If

            ' Reverse Coupled Flatness
            If SpecType = "DUAL DIRECTIONAL COUPLER" And Me.ckTest5.Checked Then
                PassFail = Tests.CoupledFlatness(2, SpecType, , TestID)
                RetrnVal = COuP
                SaveTestData("CoupledFlatness", RetrnVal)
                status("Blue", "TEST5")
                Data5.Text = Format(RetrnVal, "0.00")
                If PassFail = "Pass" Then
                    TEST5PASS = True
                    status("Green", "TEST5", True)
                    PF5.Text = PassFail
                    If TEST5Fail > 0 Then TEST5Fail = TEST5Fail - 1
                    If RetEST5Fail > 0 Then RetEST5Fail = 0
                    If Not GlobalFailed Then GlobalFail = GlobalFail - 1

                ElseIf PassFail = "Fail" Then
                    TEST5PASS = False
                    status("Red", "TEST5", True)
                    If RetEST5Fail > retestFailMax And Not TEST5Fail_bypass And Not Master_bypass And Not Retest_bypass Then
                        TEST5Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation + 50)
                        ERR.ShowDialog()
                        TEST5Failing = False
                        Retest_bypass = True
                        ErrorTimer.Stop()
                    End If
                    PF5.Text = PassFail
                    status("Red", "TEST5", True)
                    Passed = False
                    UUTFail = 1
                End If
                Me.Total5.Text = UUTNum
            Else
                'TEST5PASS = True
                'status("Blue", "TEST5")
                'PF5.Text = "Pass"
                'Me.Total5.Text = UUTNum
            End If
ReallyComplete:
            If TEST1PASS And TEST2PASS And TEST3PASS And TEST4PASS And TEST5PASS Then
                status("Green", "TEST1")
                status("Green", "TEST2")
                status("Green", "TEST3")
                status("Green", "TEST4")
                status("Green", "TEST5")
                Me.UUTStatusColor.BackColor = Color.LawnGreen
                Me.StartTestFrame.BackColor = Color.LawnGreen
                Me.ReTestFrame.BackColor = Color.LawnGreen
                If Not TweakMode Then SaveFailureLog("UUT Number: " & UUTNum & " Job Number: " & Job & " Part Number: " & Part & "  Insertion Loss: " & Me.PF1.Text & "  Return Loss: " & Me.PF2.Text & "  Coupling: " & Me.PF3.Text & "  Directivity: " & TEST4PASS & "  Coupled Balance: " & Me.PF5.Text)
                TestRunning = False
                TestComplete = True
                TestPassFail = True
                LOTFail = LOTFail - 1
                UUTFail = UUTFail - 1
                Me.LotFailure.Text = FormatPercent(((LOTFail / UUTNum)), 1)
                cmdRetest.Enabled = False
                EraseTest.Enabled = False
                If Not GlobalFailed Then
                    GlobalFail = GlobalFail + 1
                    GlobalFailed = True
                End If
            Else
                cmdRetest.Enabled = True
                EraseTest.Enabled = True
                TestRunning = False
                TestComplete = True
                TestPassFail = False
            End If
            If TEST1PASS And TEST1FailRetest >= 1 Then  ' Only update if the last run failed
                ' If Test1Fail > 0 Then Test1Fail = Test1Fail - 1
                Me.FailTotal1.Text = Test1Fail
                Me.Failures1.Text = FormatPercent(((Test1Fail / UUTNum)), 1)
            End If

            If TEST2PASS And TEST2FailRetest >= 1 Then  ' Only update if the last run failed
                'If Test2Fail > 0 Then Test2Fail = Test2Fail - 1
                Me.FailTotal2.Text = Test2Fail
                Me.Failures2.Text = FormatPercent(((Test2Fail / UUTNum)), 1)
            End If

            If TEST3PASS And TEST3FailRetest >= 1 Then  ' Only update if the last run failed
                'If TEST3Fail > 0 Then TEST3Fail = TEST3Fail - 1
                Me.FailTotal3.Text = TEST3Fail
                Me.Failures3.Text = FormatPercent(((TEST3Fail / UUTNum)), 1)
            End If

            If TEST4PASS And TEST4FailRetest >= 1 Then  ' Only update if the last run failed
                ' If TEST4Fail > 0 Then TEST4Fail = TEST4Fail - 1
                Me.FailTotal4.Text = TEST4Fail
                Me.Failures4.Text = FormatPercent(((TEST4Fail / UUTNum)), 1)
            End If
            If TEST5PASS And TEST5FailRetest >= 1 Then  ' Only update if the last run failed
                'If TEST5Fail > 0 Then TEST5Fail = TEST5Fail - 1
                Me.FailTotal5.Text = TEST5Fail
                Me.Failures5.Text = FormatPercent(((TEST5Fail / UUTNum)), 1)
            End If
            If LOTFail = 0 Then Me.LotTestFrame.BackColor = Color.LawnGreen
            If LOTFail = 0 Then Me.LotFailureFrame.BackColor = Color.LawnGreen
            If LOTFail > 0 Then Me.LotTestFrame.BackColor = Color.Red
            If LOTFail > 0 Then Me.LotFailureFrame.BackColor = Color.Red
            Me.cmdRetest.Enabled = True
            EraseTest.Enabled = True
            If UUTFail = 1 Then Me.cmdRetest.Text = "Re - Test"
            'If UUTFail = 0 Then Me.cmdRetest.Text = "Undo"

            If Passed Then Removed = RemoveLastEntryFailureLog()
            If TweakMode Then
                UUTMessage.Text = "  UUT TESTS  -- Tweak Mode. No Data Logging"
            ElseIf Not TraceChecked Then
                UUTMessage.Text = "  UUT TESTS Marker Mode  --   Load Unit #" & UUTNum + 1
            Else
                UUTMessage.Text = "  UUT TESTS  --  Load Unit #" & UUTNum + 1
            End If

            Me.cmdStartTest.Text = "Next UUT"
            cmdStartTest.Enabled = True

            MSChart.UpDateChart(SpecType)
            If LastUUT <> 0 Then
                UUTNum = LastUUT
                UUTCount.Text = LastUUT
            End If
        Else
            If UUTFail = 1 Then UndoUUT(True)
            If UUTFail = 0 Then UndoUUT(False)
        End If
        RetestReset()

    End Sub

    Private Sub cmdStartTest_Click(sender As Object, e As EventArgs) Handles cmdStartTest.Click
        Try
            Dim PassFail As String = "Pass"
            Dim SQLstr As String
            Dim SwPos As String
            Dim Dir1Failed As Boolean
            Dim TestID As Long

            Dim TestExist As Boolean
            Dim RetrnStr As String
            Dim Workstation As String = ""
            Dim Resumed As Boolean = False

            StartTime = Now()
            TestRunning = True
            TestComplete = False
            GlobalFailed = False
            Retest1Fail = 0
            Retest2Fail = 0
            RetEST3Fail = 0
            RetEST4Fail = 0
            RetEST5Fail = 0
            ArtworkRevision = txtArtwork.Text

            If SpecAB_TF Then
                Data4L.Visible = True
                Data4H.Visible = True
                Data4.Visible = False
            Else
                Data4L.Visible = False
                Data4H.Visible = False
                Data4.Visible = True
            End If
            If ISO_TF Then
                Data3L.Visible = True
                Data3H.Visible = True
                Data3.Visible = False
            Else
                Data3L.Visible = False
                Data3H.Visible = False
                Data3.Visible = True
            End If
            If IL_TF Then
                Data1L.Visible = True
                Data1H.Visible = True
                Data1.Visible = False
            Else
                Data1L.Visible = False
                Data1H.Visible = False
                Data1.Visible = True
            End If



            ResetTests()
            GetLoss()
            SwPos = ""
            If DontclickTheButton = True Then Exit Sub
            If Me.cmbJob.Text = "" Or Me.cmbJob.Text = " " Then
                MsgBox("Please select Job")
                MutiCal.Checked = False
                Exit Sub
            End If
            If Me.txtArtwork.Text = "" Then
                MsgBox("Please input the Artwork Revision")
                Exit Sub
            End If

            If Not Simulation.Checked And Not Connected Then
                ScanGPIB.connect("GPIB0::16::INSTR", GetTimeout())
                Connected = True
            End If
            If Retest1Fail > retestFailMax Then

            End If
            SwitchPorts = SQL.GetSpecification("SwitchPorts")
            ActiveTitle = Me.txtTitle.Text
            FailCount = 0
            Dir1Failed = False
            If Me.cmbJob.Text = "" Then
                MsgBox("Please Choose the Job Number first", , "Not Ready to Start")
                MutiCal.Checked = False
                ClearStatusLog()
                Exit Sub
            Else
                If cmdStartTest.Text = " Find  Offsets" Then
                    FindOffsets()
                    Exit Sub
                End If
                If Not TweakMode And (Not TEST1PASS Or Not TEST2PASS Or Not TEST3PASS Or Not TEST4PASS Or Not TEST5PASS) Then
                    If MsgBox("Are you sure you want to fail UUT" & UUTNum, vbYesNo, "Are you sure??") = vbNo Then Exit Sub
                End If
                TestExist = False
                If resumeTest Then
                    LastTest = Date.Now.Ticks
                    StatusLog.Items.Add("Testing Resumed:" & DateTime.Now.ToString)
                    PPH = SQL.GetSpecification("PartsPerHour")
                    Quantity = SQL.GetSpecification("Quantity")
                    If Me.ckROBOT.Checked Then
                        RobotTimer.Stop()
                        While GetROBOTMovingSignal()
                            'wait for robot to stop
                        End While
                        DeleteOp.Checked = False
                        Dim OP As New OperatorEntry
                        OP.StartPosition = FormStartPosition.Manual
                        OP.Location = New Point(globals.XLocation, globals.YLocation)
                        OP.ShowDialog()
                        TestCompleteSignal(False)
                        RobotTimer.Start()
                    Else
                        DeleteOp.Checked = False
                        Dim OP As New OperatorEntry
                        OP.StartPosition = FormStartPosition.Manual
                        OP.Location = New Point(globals.XLocation, globals.YLocation)
                        OP.ShowDialog()
                    End If
                    'ATS Developer does not save data while developing
                    If User = "ATS" Then
                        GetTrace.CheckState = CheckState.Unchecked
                    End If
                    UUTCount.Text = UUTNum
                    StarStartExpectedTimeline()
                    resumeTest = False
                    EndLot.Enabled = True
                    Me.EraseJob.Enabled = True
                    Resumed = True
                    OperatorLog = True
                End If
            End If
            If UUTNum = 0 And Not Resumed Then
                ResetTests()
                stopTest = False
                LastTest = Date.Now.Ticks
                ClearStatusLog()
                StatusLog.Items.Add("Testing Started:" & DateTime.Now.ToString)
                PPH = SQL.GetSpecification("PartsPerHour")
                Quantity = SQL.GetSpecification("Quantity")
                If Me.ckROBOT.Checked Then
                    RobotTimer.Stop()
                    While Not GetROBOTMovingSignal()
                        'wait for robot to stop
                    End While
                    DeleteOp.Checked = False
                    Dim OP As New OperatorEntry
                    OP.StartPosition = FormStartPosition.Manual
                    OP.Location = New Point(globals.XLocation, globals.YLocation)
                    OP.ShowDialog()
                    TestCompleteSignal(False)
                    RobotTimer.Start()
                Else
                    DeleteOp.Checked = False
                    Dim OP As New OperatorEntry
                    OP.StartPosition = FormStartPosition.Manual
                    OP.Location = New Point(globals.XLocation, globals.YLocation)
                    OP.ShowDialog()
                End If
                UUTCount.Text = UUTNum
                StarStartExpectedTimeline()
                EndLot.Enabled = True
                Me.EraseJob.Enabled = True
                OperatorLog = True
            ElseIf Resumed Then
                UUTCount.Text = UUTNum
                StarStartExpectedTimeline()
                EndLot.Enabled = True
                Me.EraseJob.Enabled = True
                OperatorLog = True
            End If

            UUTFail = 0
            Me.RunStatus.ForeColor = Color.Red
            Data1.Text = ""
            Data1L.Text = ""
            Data1H.Text = ""
            Data2.Text = ""
            Data3.Text = ""
            Data4.Text = ""
            Data4L.Text = ""
            Data4H.Text = ""
            Data5.Text = ""
            If UUTNum = 0 Then UUTNum = UUTNum + 1

            If (UUTNum = 1 Or FirstPart) And stopTest Then
                ResetLot()
                Exit Sub
            End If
            FirstPart = False
            If UUTNum = 1 Or UUTReset Then
                UUTNum_Reset = 1
                UUTReset = False
                Me.GetTrace.Checked = True
            End If

            ReportServer("test running", UUTNum, False)
            saveConfigurationVal(iniPathName, "uut_number", UUTNum)
            If TweakMode Then UUTNum = 1
            If UUTNum_Reset > 5 And Not BypassUnchecked Then
                Me.GetTrace.Checked = False
            ElseIf UUTNum_Reset > 5 And BypassUnchecked Then
                Me.GetTrace.Checked = True
            End If

            UUTCount.Text = Str(UUTNum)
            SQL.UpdateEffeciency("Running", txtEfficiency.Text, Now.Date.ToShortDateString, UUTCount.Text)
            If TweakMode Then
                Me.UUTMessage.Text = "  UUT TESTS  --   Testing Undisclosed Unit "
            ElseIf Not TraceChecked Then
                UUTMessage.Text = "  UUT TESTS Marker Mode  --   Load Unit #" & UUTNum + 1
            Else
                Me.UUTMessage.Text = "  UUT TESTS  --   Testing Unit #" & UUTNum
            End If

            Me.cmdStartTest.Text = "UUT" & UUTNum
            Me.cmdStartTest.Enabled = False
            Me.Refresh()

            If Me.cmbJob.Text = " " Then Exit Sub
            Job = Me.cmbJob.Text
            Part = Me.cmbPart.Text
            SerialNumber = "UUT" & UUTNum_Reset
            SQLstr = "SELECT * from TestData where JobNumber = '" & Job & "' And SerialNumber = '" & SerialNumber & "' and WorkStation = '" & GetComputerName() & "' and artwork_rev = '" & ArtworkRevision & "'"
            If SQL.CheckforRow(SQLstr, "NetworkData") = 0 Then
                SQLstr = "Insert Into TestData (JobNumber, PartNumber,SerialNumber,WorkStation,artwork_rev) values ('" & Job & "','" & Part & "','" & SerialNumber & "','" & GetComputerName() & "','" & ArtworkRevision & "')"
                SQL.ExecuteSQLCommand(SQLstr, "NetworkData")
            End If
            SQLstr = "SELECT * from TestData where JobNumber = '" & Job & "' And SerialNumber = '" & SerialNumber & "' and WorkStation = '" & GetComputerName() & "' and artwork_rev = '" & ArtworkRevision & "'"
            TestID = SQL.GetTestID(SQLstr, "NetworkData")
            Me.Refresh()

            'Insertion Loss
            If Me.ckTest1.Checked Then
                If Me.ckROBOT.Checked Then RobotStatus()
                If TraceChecked Then
                    If (SpecType = "TRANSFORMER") And IL_TF Then PassFail = Tests.InsertionLossTRANS_multiband(, TestID)
                    If SpecType = "TRANSFORMER" Then PassFail = Tests.InsertionLossTRANS(, TestID)
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN") And SpecAB_TF Then PassFail = Tests.InsertionLoss3dB_multiband(, TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.InsertionLoss3dB(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.InsertionLossCOMB(, TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.InsertionLossDIR(, TestID)
                ElseIf Not MutiCalChecked Then
                    If (SpecType = "TRANSFORMER") And IL_TF Then PassFail = Tests.InsertionLossTRANS_multiband(, TestID)
                    If SpecType = "TRANSFORMER" Then PassFail = Tests.InsertionLossTRANS_Marker(, TestID)
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And SpecAB_TF Then PassFail = Tests.InsertionLoss3dB_multiband(, TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.InsertionLoss3dB_marker(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.InsertionLossCOMB_Marker(, TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.InsertionLossDIR_Marker(, TestID)
                End If


                If IL_TF Then
                    RetrnVal1 = IL1
                    SaveTestData("InsertionLoss", RetrnVal1)
                    RetrnVal = IL2
                    SaveTestData("InsertionLoss2", RetrnVal)
                    Data1L.Text = IL1
                    Data1H.Text = IL2
                Else
                    RetrnVal = IL
                    SaveTestData("InsertionLoss", RetrnVal)
                    RetrnStr = CStr(TruncateDecimal(RetrnVal, 2))
                    Data1.Text = Format(RetrnVal, "0.00")
                End If
                status("Blue", "TEST1")
                PF1.Text = PassFail


                If PassFail = "Pass" Then
                    TEST1PASS = True
                    status("Green", "TEST1")
                    MSChart.UpDateChartData(SpecType, "IL", "Pass")
                ElseIf PassFail = "Fail" Then
                    TEST1PASS = False
                    status("Red", "TEST1")
                    Test1Fail = Test1Fail + 1
                    If Not GlobalFailed Then
                        GlobalFail = GlobalFail + 1
                        GlobalFailed = True
                    End If
                    If GlobalFail > GlobalFailMax And Not GlobalFail_bypass And Not Master_bypass Then
                        GlobalFailing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation)
                        ERR.ShowDialog()
                        ErrorTimer.Stop()
                    End If
                    If Test1Fail > TestFailMax And Not Test1Fail_bypass And Not Master_bypass Then
                        Test1Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation)
                        ERR.ShowDialog()
                        Test1Failing = False
                        ErrorTimer.Stop()
                    End If
                    status("Red", "TEST1")
                    TEST1FailRetest = TEST1FailRetest + 1
                    UUTFail = 1
                End If
                Me.Failures1.Text = FormatPercent(((Test1Fail / UUTNum)), 1)
                Me.Total1.Text = UUTNum
                Me.FailTotal1.Text = Test1Fail
            Else
                TEST1PASS = True
                status("Blue", "TEST1")
                SaveTestData("InsertionLoss", GetSpecification("InsertionLoss"))
                Me.Failures1.Text = FormatPercent(((Test1Fail / UUTNum)), 1)
                Me.Total1.Text = UUTNum
                Me.FailTotal1.Text = Test1Fail
            End If
            Me.Refresh()

            'Return Loss
            If Me.ckTest2.Checked Then
                If Me.ckROBOT.Checked Then RobotStatus()
                If TraceChecked And Not TweakMode Then
                    PassFail = Tests.ReturnLoss(, TestID)
                Else
                    PassFail = Tests.ReturnLoss_Marker(, TestID)
                End If
                RetrnVal = RL
                SaveTestData("ReturnLoss", RetrnVal)
                status("Blue", "TEST2")
                PF2.Text = PassFail
                RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                Data2.Text = Format(RetrnVal, "0.0")
                If PassFail = "Pass" Then
                    TEST2PASS = True
                    status("Green", "TEST2")
                ElseIf PassFail = "Fail" Then
                    TEST2PASS = False
                    status("Red", "TEST2")
                    Test2Fail = Test2Fail + 1
                    If Not GlobalFailed Then
                        GlobalFail = GlobalFail + 1
                        GlobalFailed = True
                    End If
                    If GlobalFail > GlobalFailMax And Not GlobalFail_bypass And Not Master_bypass Then
                        GlobalFailing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation)
                        ERR.ShowDialog()
                        ErrorTimer.Stop()
                    End If
                    If Test2Fail > TestFailMax And Not Test2Fail_bypass And Not Master_bypass Then
                        Test2Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation)
                        ERR.ShowDialog()
                        Test2Failing = False
                        ErrorTimer.Stop()
                    End If

                    status("Red", "TEST2")
                    TEST2FailRetest = TEST2FailRetest + 1
                    UUTFail = 1
                End If
                Me.Failures2.Text = FormatPercent(((Test2Fail / UUTNum)), 1)
                Me.Total2.Text = UUTNum
                Me.FailTotal2.Text = Test2Fail
            Else
                TEST2PASS = True
                status("Blue", "TEST2")
                SaveTestData("ReturnLoss", VSWRtoRL(SQL.GetSpecification("VSWR")))
                Me.Failures2.Text = FormatPercent(((Test2Fail / UUTNum)), 1)
                Me.Total2.Text = UUTNum
                Me.FailTotal2.Text = Test2Fail
            End If
            Me.Refresh()
            If SpecType <> "COMBINER/DIVIDER" Then GoTo Test2Sub
Test2SubRet:

            'AmplitudeBalance
            'Directivity
            If Me.ckTest4.Checked Then
                If Me.ckROBOT.Checked Then RobotStatus()
                If TraceChecked Then
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And SpecAB_TF Then PassFail = Tests.AmplitudeBalance_multiband(, TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.AmplitudeBalance(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.AmplitudeBalanceCOMB(, TestID)
                    If SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Then PassFail = Tests.Directivity(1, SpecType, , TestID)
                ElseIf Not MutiCalChecked Then
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And SpecAB_TF Then PassFail = Tests.AmplitudeBalance_multiband(, TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.AmplitudeBalance_Marker(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.AmplitudeBalanceCOMB_Marker(, TestID)
                    If SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Then PassFail = Tests.Directivity_Marker(1, SpecType, , TestID)
                End If
                If SpecType <> "SINGLE DIRECTIONAL COUPLER" Then
                    If SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Then
                        RetrnVal = DIR
                        SaveTestData("Directivity", RetrnVal)
                    Else
                        If SpecAB_TF Then
                            RetrnVal = AB1
                            SaveTestData("AmplitudeBalance1", RetrnVal)
                            RetrnVal = AB2
                            SaveTestData("AmplitudeBalance2", RetrnVal)
                            'remove later
                            RetrnVal = AB
                            SaveTestData("AmplitudeBalance", RetrnVal)
                        Else
                            RetrnVal = AB
                            SaveTestData("AmplitudeBalance", RetrnVal)
                        End If
                    End If
                    status("Blue", "TEST4")
                    If SpecAB_TF Then
                        AB1 = Format(AB1, "0.00")
                        AB2 = Format(AB2, "0.00")
                        Data4L.Text = AB1
                        Data4H.Text = AB2
                        If AB1Pass = "Pass" And AB2Pass = "Pass" Then
                            TEST4PASS = True
                            PF4.Text = PassFail
                            status("Green", "TEST4L", True)
                            status("Green", "TEST4H", True)
                            If SpecType <> "DUAL DIRECTIONAL COUPLER" And SpecType <> "SINGLE DIRECTIONAL COUPLER" Then
                                TEST4PASS = True
                            End If
                        ElseIf AB1Pass = "Fail" Or AB2Pass = "Fail" Then
                            TEST4PASS = False
                            PassFail = "Fail"
                            PF4.Text = PassFail
                            If AB1Pass = "Pass" Then
                                status("Green", "TEST4L", True)
                            ElseIf AB1Pass = "Fail" Then
                                status("Red", "TEST4L", True)
                            End If
                            If AB2Pass = "Pass" Then
                                status("Green", "TEST4H", True)
                            ElseIf AB2Pass = "Fail" Then
                                status("Red", "TEST4H", True)
                            End If
                            If SpecType <> "DUAL DIRECTIONAL COUPLER" And SpecType <> "SINGLE DIRECTIONAL COUPLER" Then
                                TEST4PASS = False
                                TEST4Fail = TEST4Fail + 1
                                If Not GlobalFailed Then
                                    GlobalFail = GlobalFail + 1
                                    GlobalFailed = True
                                End If
                                If GlobalFail > GlobalFailMax And Not GlobalFail_bypass And Not Master_bypass Then
                                    GlobalFailing = True
                                    ErrorTimer.Start()
                                    Dim ERR As New Failures_Manager
                                    ERR.StartPosition = FormStartPosition.Manual
                                    ERR.ShowDialog()
                                    ErrorTimer.Stop()
                                End If
                                If TEST4Fail > TestFailMax And Not TEST4Fail_bypass And Not Master_bypass Then
                                    TEST4Failing = True
                                    ErrorTimer.Start()
                                    Dim ERR As New Failures_Manager
                                    ERR.StartPosition = FormStartPosition.Manual
                                    ERR.Location = New Point(globals.XLocation, globals.YLocation)
                                    ERR.ShowDialog()
                                    TEST4Failing = False
                                    ErrorTimer.Stop()
                                End If
                                status("Red", "TEST4")
                                TEST4FailRetest = TEST4FailRetest + 1
                                UUTFail = 1
                            End If
                        End If
                    Else
                        If PassFail = "Pass" Then
                            status("Green", "TEST4")
                            Data4.Text = Format(RetrnVal, "0.00")
                            PF4.Text = PassFail
                            TEST4PASS = True
                        ElseIf PassFail = "Fail" Then
                            status("Red", "TEST4")
                            Data4.Text = Format(RetrnVal, "0.00")
                            PF4.Text = PassFail
                            TEST4PASS = False
                            TEST4Fail = TEST4Fail + 1
                            Me.Failures4.Text = FormatPercent(((TEST4Fail / UUTNum)), 1)
                            Me.Total4.Text = UUTNum
                            Me.FailTotal4.Text = TEST4Fail
                            If Not GlobalFailed Then
                                GlobalFail = GlobalFail + 1
                                GlobalFailed = True
                            End If
                            If GlobalFail > GlobalFailMax And Not GlobalFail_bypass And Not Master_bypass Then
                                GlobalFailing = True
                                ErrorTimer.Start()
                                Dim ERR As New Failures_Manager
                                ERR.StartPosition = FormStartPosition.Manual
                                ERR.ShowDialog()
                                ErrorTimer.Stop()
                            End If
                            If TEST4Fail > TestFailMax And Not TEST4Fail_bypass And Not Master_bypass Then
                                TEST4Failing = True
                                ErrorTimer.Start()
                                Dim ERR As New Failures_Manager
                                ERR.StartPosition = FormStartPosition.Manual
                                ERR.Location = New Point(globals.XLocation, globals.YLocation)
                                ERR.ShowDialog()
                                TEST4Failing = False
                                ErrorTimer.Stop()
                            End If
                            status("Red", "TEST4")
                            TEST4FailRetest = TEST4FailRetest + 1
                            UUTFail = 1
                        End If
                    End If
                End If
            End If
            Me.Refresh()
            'PhaseBalance
            'CoupledFlatness
            If Me.ckTest5.Checked Then
                If Me.ckROBOT.Checked Then RobotStatus()
                If TraceChecked Then
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.PhaseBalance(SpecType, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.PhaseBalanceCOMB(SpecType, , TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.CoupledFlatness(1, SpecType, , TestID)
                ElseIf Not MutiCalChecked Then
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.PhaseBalance_Marker(SpecType, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.PhaseBalanceCOMB_Marker(SpecType, , TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.CoupledFlatness(1, SpecType, , TestID)
                End If
                If InStr(SpecType, "DIRECTIONAL COUPLER") Then
                    RetrnVal = CF
                    SaveTestData("CoupledFlatness", RetrnVal)
                Else
                    RetrnVal = PB
                    SaveTestData("PhaseBalance", RetrnVal)
                End If
                status("Blue", "TEST5")
                PF5.Text = PassFail

                RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                Data5.Text = Format(RetrnVal, "0.0")
                If PassFail = "Pass" Then
                    status("Green", "TEST5")
                    If SpecType <> "DUAL DIRECTIONAL COUPLER" Then
                        TEST5PASS = True
                    End If
                ElseIf PassFail = "Fail" Then
                    status("Red", "TEST5")
                    If SpecType <> "DUAL DIRECTIONAL COUPLER" Then
                        TEST5PASS = False
                        TEST5Fail = TEST5Fail + 1
                        If Not GlobalFailed Then
                            GlobalFail = GlobalFail + 1
                            GlobalFailed = True
                        End If
                        If GlobalFail > GlobalFailMax And Not GlobalFail_bypass And Not Master_bypass Then
                            GlobalFailing = True
                            ErrorTimer.Start()
                            Dim ERR As New Failures_Manager
                            ERR.StartPosition = FormStartPosition.Manual
                            ERR.Location = New Point(globals.XLocation, globals.YLocation)
                            ERR.ShowDialog()
                            ErrorTimer.Stop()
                        End If
                        If TEST5Fail > TestFailMax And Not TEST5Fail_bypass And Not Master_bypass Then
                            TEST5Failing = True
                            ErrorTimer.Start()
                            Dim ERR As New Failures_Manager
                            ERR.StartPosition = FormStartPosition.Manual
                            ERR.Location = New Point(globals.XLocation, globals.YLocation)
                            ERR.ShowDialog()
                            TEST5Failing = False
                            ErrorTimer.Stop()
                        End If
                        status("Red", "TEST5")
                        TEST5FailRetest = TEST5FailRetest + 1
                        UUTFail = 1
                    End If
                End If
                If SpecType <> "DUAL DIRECTIONAL COUPLER" Then
                    Me.Failures5.Text = FormatPercent(((TEST5Fail / UUTNum)), 1)
                    Me.Total5.Text = UUTNum
                    Me.FailTotal5.Text = TEST5Fail
                End If
            Else
                TEST5PASS = True
                status("Blue", "TEST5")
                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then SaveTestData("PhaseBalance", GetSpecification("PhaseBalance"))
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then SaveTestData("CoupledFlatness", GetSpecification("CoupledFlatness"))

                Me.Failures5.Text = FormatPercent(((TEST5Fail / UUTNum)), 1)
                Me.Total5.Text = UUTNum
                Me.FailTotal5.Text = TEST5Fail
            End If

            If SpecType <> "COMBINER/DIVIDER" Then
                If SpecType = "DUAL DIRECTIONAL COUPLER" And PF1.Text <> "Fail" And PF2.Text <> "Fail" And (PF3.Text = "Fail" Or PF4.Text = "Fail" Or PF5.Text = "Fail") Then
                    If MsgBox("Try Forward Measurement Again", vbYesNo) = vbYes Then
                        GoTo Test2Sub
                    Else
                        GoTo TestComplete
                    End If
                    Dir1Failed = True
                End If
                GoTo TestComplete
            End If
            Me.Refresh()
Test2Sub:

            'Isolation
            'Coupling
            If Me.ckTest3.Checked Then
                If Me.ckROBOT.Checked Then RobotStatus()
                If TraceChecked Then
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.Isolation(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.IsolationCOMB(, TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.Coupling(1, SpecType, , TestID)
                ElseIf Not MutiCalChecked Then
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.Isolation_Marker(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") And ISO_TF Then PassFail = Tests.IsolationCOMB(, TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.IsolationCOMB_Marker(, TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.Coupling(1, SpecType, , TestID)
                End If
                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Or SpecType = "SINGLE DIRECTIONAL COUPLER" Then
                    If SpecType.Contains("COMBINER/DIVIDER") And ISO_TF Then
                        ISoL = Format(ISoL, "0.00")
                        ISoH = Format(ISoH, "0.00")
                        Data3L.Text = ISoL
                        Data3H.Text = ISoH
                        PF3.Text = PassFail
                        If PassFail = "Pass" Then
                            TEST3PASS = True
                            status("Green", "TEST3L", True)
                            status("Green", "TEST3H", True)
                        ElseIf PassFail = "Fail" Then
                            TEST3PASS = False
                            If AB1Pass = "Pass" Then
                                status("Green", "TEST3L", True)
                            ElseIf AB1Pass = "Fail" Then
                                status("Red", "TEST3L", True)
                            End If
                            If AB2Pass = "Pass" Then
                                status("Green", "TEST3H", True)
                            ElseIf AB2Pass = "Fail" Then
                                status("Red", "TEST3H", True)
                            End If
                        End If

                        If SQLAccess Then
                            SaveTestData("IsolationL", ISoL)
                            RetrnStr = CStr(TruncateDecimal(ISoL, 1))
                            SaveTestData("IsolationH", ISoH)
                            RetrnStr = CStr(TruncateDecimal(ISoH, 1))
                        Else
                            SaveTestData("IsoL", ISoL)
                            RetrnStr = CStr(TruncateDecimal(ISoL, 1))
                            SaveTestData("IsoH", ISoH)
                            RetrnStr = CStr(TruncateDecimal(ISoH, 1))
                        End If
                    ElseIf SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("COMBINER/DIVIDER") Or SpecType.Contains("BALUN") Then
                        RetrnVal = ISo
                        If SQLAccess Then
                            SaveTestData("Isolation", RetrnVal)
                            RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                        Else
                            SaveTestData("Iso", RetrnVal)
                            RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                        End If
                    Else
                        RetrnVal = COuP
                        SaveTestData("Coupling", RetrnVal)
                        RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                    End If
                    If Not ISO_TF Then
                        status("Blue", "TEST3")
                        PF3.Text = PassFail
                        Data3.Text = Format(RetrnVal, "0.0")
                        If PassFail = "Pass" Then
                            status("Green", "TEST3")
                            If SpecType <> "DUAL DIRECTIONAL COUPLER" Then
                                TEST3PASS = True
                            End If
                        ElseIf PassFail = "Fail" Then
                            status("Red", "TEST3")
                            If SpecType <> "DUAL DIRECTIONAL COUPLER" Then
                                TEST3PASS = False
                                TEST3Fail = TEST3Fail + 1
                                If Not GlobalFailed Then
                                    GlobalFail = GlobalFail + 1
                                    GlobalFailed = True
                                End If
                                If GlobalFail > GlobalFailMax And Not GlobalFail_bypass And Not Master_bypass Then
                                    GlobalFailing = True
                                    ErrorTimer.Start()
                                    Dim ERR As New Failures_Manager
                                    ERR.StartPosition = FormStartPosition.Manual
                                    ERR.Location = New Point(globals.XLocation, globals.YLocation)
                                    ERR.ShowDialog()
                                    ErrorTimer.Stop()
                                End If
                                If TEST3Fail > TestFailMax And Not TEST3Fail_bypass And Not Master_bypass Then
                                    TEST3Failing = True
                                    ErrorTimer.Start()
                                    Dim ERR As New Failures_Manager
                                    ERR.StartPosition = FormStartPosition.Manual
                                    ERR.Location = New Point(globals.XLocation, globals.YLocation)
                                    ERR.ShowDialog()
                                    TEST3Failing = False
                                    ErrorTimer.Stop()
                                End If
                                status("Red", "TEST3")
                                TEST3FailRetest = TEST3FailRetest + 1
                                UUTFail = 1
                            End If
                        End If
                    End If
                    If SpecType <> "DUAL DIRECTIONAL COUPLER" Then
                        Me.Failures3.Text = FormatPercent(((TEST3Fail / UUTNum)), 1)
                        Me.Total3.Text = UUTNum
                        Me.FailTotal3.Text = TEST3Fail
                    End If
                Else
                    TEST3PASS = True
                    status("Blue", "TEST3")
                    If SQLAccess Then
                        If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") And SQLAccess Then SaveTestData("Isolation", 0 - GetSpecification("Isolation"))
                        If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") And Not SQLAccess Then SaveTestData("Isolation", 0 - GetSpecification("Isolation"))
                    Else
                        If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") And SQLAccess Then SaveTestData("Iso", 0 - GetSpecification("Isolation"))
                        If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") And Not SQLAccess Then SaveTestData("Iso", 0 - GetSpecification("Isolation"))
                    End If

                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then SaveTestData("Coupling", 0 - GetSpecification("Coupling"))
                    Me.Failures3.Text = FormatPercent(((TEST3Fail / UUTNum)), 1)
                    Me.Total3.Text = UUTNum
                    Me.FailTotal3.Text = TEST3Fail
                End If
                Me.Refresh()
            End If

            If Not SpecType.Contains("COMBINER/DIVIDER") Then GoTo Test2SubRet

TestComplete:  ' For everything except Directional Couplers

            'Directonal Couplers reverse direction
            If Not TweakMode And (SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or (SpecType = "SINGLE DIRECTIONAL COUPLER" And Me.ckTest4.Checked)) Then
                If SpecType = "DUAL DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, REFL = SW3"
                If SpecType = "DUAL DIRECTIONAL COUPLER" And SwitchPorts = 1 Then SwPos = " OUT = SW1, CPL_J4 = SW2, ISO_J3 = SW3, CPL_J3 = SW4, ISO_J4 = SW5"
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, ISO = SW3"
                If SpecType = "BI DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, REFL = SW3"
                txtTitle.Text = SpecType & SwPos
                MsgBox("Please turn the Directional Coupler in the Reverse direction")
            End If

            'Reverse Coupling
            If SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" And Me.ckTest3.Checked Then
                If Me.ckROBOT.Checked Then RobotStatus()
                PassFail = Tests.Coupling(2, SpecType, , TestID)
                RetrnVal = COuP
                RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                SaveTestData("Coupling", RetrnVal)
                status("Blue", "TEST3")
                PF3.Text = PassFail
                Data3.Text = Format(RetrnVal, "0.00")
                If PassFail = "Pass" Then
                    TEST3PASS = True
                    status("Green", "TEST3")
                ElseIf PassFail = "Fail" Then
                    TEST3PASS = False
                    status("Red", "TEST3")
                    TEST3Fail = TEST3Fail + 1
                    If Not GlobalFailed Then
                        GlobalFail = GlobalFail + 1
                        GlobalFailed = True
                    End If
                    If GlobalFail > GlobalFailMax And Not GlobalFail_bypass And Not Master_bypass Then
                        GlobalFailing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation)
                        ERR.ShowDialog()
                        ErrorTimer.Stop()
                    End If
                    status("Red", "TEST3")
                    TEST3FailRetest = TEST3FailRetest + 1
                    UUTFail = 1
                    If TEST3Fail > TestFailMax And Not TEST3Fail_bypass And Not GlobalFail_bypass And Not Master_bypass Then
                        TEST3Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation)
                        ERR.ShowDialog()
                        TEST3Failing = False
                        ErrorTimer.Stop()
                    End If
                End If
                Me.Failures3.Text = FormatPercent(((TEST3Fail / UUTNum)), 1)
                Me.Total3.Text = UUTNum
                Me.FailTotal3.Text = TEST3Fail
            End If
            Me.Refresh()

            ' Reverse Directivity
            If (SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "SINGLE DIRECTIONAL COUPLER") And Me.ckTest4.Checked Then
                If Me.ckROBOT.Checked Then RobotStatus()
                If TraceChecked And Not TweakMode Then
                    PassFail = Tests.Directivity(2, SpecType, , TestID)
                Else
                    PassFail = Tests.Directivity_Marker(2, SpecType, , TestID)
                End If

                status("Blue", "TEST4")
                PF4.Text = PassFail
                RetrnVal = DIR
                SaveTestData("Directivity", RetrnVal)
                RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                Data4.Text = Format(RetrnVal, "0.0")
                If PassFail = "Pass" Then
                    TEST4PASS = True
                    status("Green", "TEST4")
                    MSChart.UpDateChartData(SpecType, "DIR", "Fail")
                ElseIf PassFail = "Fail" Then
                    TEST4PASS = False
                    status("Red", "TEST4")
                    TEST4Fail = TEST4Fail + 1
                    If Not GlobalFailed Then
                        GlobalFail = GlobalFail + 1
                        GlobalFailed = True
                    End If
                    If GlobalFail > GlobalFailMax And Not GlobalFail_bypass And Not Master_bypass Then
                        GlobalFailing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation)
                        ERR.ShowDialog()
                        ErrorTimer.Stop()
                    End If
                    status("Red", "TEST4")
                    TEST4FailRetest = TEST4FailRetest + 1
                    MSChart.UpDateChartData(SpecType, "DIR", "Pass")
                    UUTFail = 1
                    If TEST4Fail > TestFailMax And Not TEST4Fail_bypass And Not Master_bypass Then
                        TEST4Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation)
                        ERR.ShowDialog()
                        TEST4Failing = False
                        ErrorTimer.Stop()
                    End If
                End If
                Me.Failures4.Text = FormatPercent(((TEST4Fail / UUTNum)), 1)
                Me.Total4.Text = UUTNum
                Me.FailTotal4.Text = TEST4Fail
            End If
            Me.Refresh()
            ' Reverse Coupled Flatness
            If SpecType = "DUAL DIRECTIONAL COUPLER" And Me.ckTest5.Checked Then
                If Me.ckROBOT.Checked Then RobotStatus()
                PassFail = Tests.CoupledFlatness(2, SpecType, , TestID)
                RetrnVal = CF
                SaveTestData("CoupledFlatness", RetrnVal)
                status("Blue", "TEST5")
                PF5.Text = PassFail
                RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                Data5.Text = Format(RetrnVal, "0.0")
                If PassFail = "Pass" Then
                    TEST5PASS = True
                    status("Green", "TEST5")
                    MSChart.UpDateChartData(SpecType, "CB", "Pass")
                ElseIf PassFail = "Fail" Then
                    TEST5PASS = False
                    status("Red", "TEST5")
                    TEST5Fail = TEST5Fail + 1
                    If Not GlobalFailed Then
                        GlobalFail = GlobalFail + 1
                        GlobalFailed = True
                    End If
                    If GlobalFail > GlobalFailMax And Not GlobalFail_bypass And Not Master_bypass Then
                        GlobalFailing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation)
                        ERR.ShowDialog()
                        ErrorTimer.Stop()
                    End If
                    If TEST5Fail > TestFailMax And Not TEST5Fail_bypass And Not Master_bypass Then
                        TEST5Failing = True
                        ErrorTimer.Start()
                        Dim ERR As New Failures_Manager
                        ERR.StartPosition = FormStartPosition.Manual
                        ERR.Location = New Point(globals.XLocation, globals.YLocation)
                        ERR.ShowDialog()
                        TEST5Failing = False
                        ErrorTimer.Stop()
                    End If
                    status("Red", "TEST5")
                    TEST5FailRetest = TEST5FailRetest + 1
                    MSChart.UpDateChartData(SpecType, "CB", "Fail")
                    UUTFail = 1
                End If
                Me.Failures5.Text = FormatPercent(((TEST5Fail / UUTNum)), 1)
                Me.Total5.Text = UUTNum
                Me.FailTotal5.Text = TEST5Fail
            End If
            If SpecType = "DUAL DIRECTIONAL COUPLER" And Not Dir1Failed And PF1.Text <> "Fail" And PF2.Text <> "Fail" And (PF3.Text = "Fail" Or PF4.Text = "Fail" Or PF5.Text = "Fail") Then
                If MsgBox("Try Reverse Measurement Again?", vbYesNo) = vbYes Then
                    GoTo TestComplete
                End If

            End If

            If Not TweakMode And (Not TEST2PASS Or Not TEST1PASS Or Not TEST3PASS Or Not TEST4PASS Or Not TEST5PASS) Then
                UpdateFailureLog(Me.PF1.Text, Me.PF2.Text, Me.PF3.Text, Me.PF4.Text, Me.PF5.Text)
                SaveFailureLog("UUT Number: " & UUTNum & " Job Number: " & Me.cmbJob.Text & " Part Number: " & Me.cmbPart.Text & "  Insertion Loss: " & Me.PF1.Text & "  Return Loss: " & Me.PF2.Text & "  Coupling: " & Me.PF3.Text & "  Directivity: " & TEST4PASS & "  Coupled Balance: " & Me.PF5.Text)
                LOTFail = LOTFail + 1
            End If

            If LOTFail = 0 Then LotTestFrame.BackColor = Color.LawnGreen
            If LOTFail = 0 Then LotFailureFrame.BackColor = Color.LawnGreen
            If LOTFail > 0 Then LotTestFrame.BackColor = Color.Red
            If LOTFail > 0 Then LotFailureFrame.BackColor = Color.Red
            Me.cmdRetest.Enabled = True
            Me.EraseTest.Enabled = True
            If UUTFail = 1 Then Me.cmdRetest.Text = "Re - Test"
            ' If UUTFail = 0 Then Me.cmdRetest.Text = "Undo"
            If ExpectedProgress2 > 0 Then SQL.UpdateEffeciency("Running", txtEfficiency.Text, txtCurrentTime.Text, CInt(UUTCount.Text))
            Me.LotFailure.Text = FormatPercent(((LOTFail / UUTNum)), 1)
            If TweakMode Then
                UUTMessage.Text = "  UUT TESTS  --   Tweak Mode. No Data Logging"
            ElseIf Not TraceChecked Then
                UUTMessage.Text = "  UUT TESTS Marker Mode  --   Load Unit #" & UUTNum + 1
            Else
                UUTMessage.Text = "  UUT TESTS  --   Load Unit #" & UUTNum + 1
            End If
TestReallyComplete:
            Me.cmdStartTest.Text = "Next UUT"
            cmdStartTest.Enabled = True
            'MSChart.UpDateChart(SpecType)
            CheckTestResume()
            Me.Refresh()
            If Not TweakMode Then If ActualProgress.Value + 1 <= ActualProgress.Maximum Then ActualProgress.Value = ActualProgress.Value + 1
            'Change Title Back To Forward
            If SpecType = "DUAL DIRECTIONAL COUPLER" Or (SpecType = "SINGLE DIRECTIONAL COUPLER" And Me.ckTest4.Checked) Then
                If SpecType = "DUAL DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, REFL = SW3"
                If SpecType = "DUAL DIRECTIONAL COUPLER" And SwitchPorts = 1 Then SwPos = " OUT = SW1, CPL_J4 = SW2, ISO_J3 = SW3, CPL_J3 = SW4, ISO_J4 = SW5"
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, ISO = SW3"
                If SpecType = "BI DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, REFL = SW3"
                txtTitle.Text = SpecType & SwPos
            End If
            If TEST1PASS And TEST2PASS And TEST3PASS And TEST4PASS And TEST5PASS Then
                TestRunning = False
                TestComplete = True
                TestPassFail = True
            Else
                TestRunning = False
                PassFail = False
                TestPassFail = False
            End If
            UUTNum = UUTNum + 1
            UUTNum_Reset = UUTNum_Reset + 1
            If UUTNum > 5 And Not BypassUnchecked Then
                Me.GetTrace.Checked = False
            ElseIf UUTNum > 5 And BypassUnchecked Then
                Me.GetTrace.Checked = True
            End If
            TestCompleteSignal(True) ' Note False/False tells the Robot to continue
            StopTime = Now()
            TestTime = StartTime - StopTime
            EraseTest.Text = "Remove UUT" & UUTNum - 1

            ' Me.MasterUpLoad.Enabled = True
        Catch ex As Exception

        End Try
    End Sub

    Private Sub EmptyTraceData_Click()
        Dim SQLstr As String

        If MsgBox("Are you sure you want to erase All Local Trace Data", vbYesNo, "Cannot be undone.") = vbYes Then
            SQLstr = "Delete from Trace "
            SQL.ExecuteSQLCommand(SQLstr, "LocalTraceData")

            SQLstr = "Delete from TracePoints "
            SQL.ExecuteSQLCommand(SQLstr, "LocalTraceData")
        End If
    End Sub
    Private Sub EraseJob_Click(sender As Object, e As EventArgs) Handles EraseJob.Click
        Dim mes As String
        Dim SQLstr As String
        StatusLog.Items.Add("Closing the Job:" & DateTime.Now.ToString)
        ReportServer("job closed", UUTCount.Text, False)
        SQL.UpdateEffeciency("job closed", txtEfficiency.Text, txtCurrentTime.Text, UUTCount.Text)
        txtCurrentTime.Text = "CLOSED"
        ExpectedTimer.Stop()
        ClearStatusLog()
        StatusLog.Items.Add("Closed" & "   " & DateTime.Now.ToString)
        mes = "Job has been closed for this workstation by operator"
        StatusLog.Items.Add(mes)
        ResetLot()
        cmbJob.Text = " "
        FirstPart = False
        If MsgBox("Do you want to erase the data from This Job", vbYesNo, "Cannot be undone.") = vbYes Then
            EraseThisTest()
        End If
        SQLstr = "Delete from Trace "
        SQL.ExecuteSQLCommand(SQLstr, "LocalTraceData")

        SQLstr = "Delete from TracePoints "
        SQL.ExecuteSQLCommand(SQLstr, "LocalTraceData")
    End Sub

    Private Sub EraseTest_Click(sender As Object, e As EventArgs) Handles EraseTest.Click
        Dim mes As String
        UUTNum = UUTNum - 1
        UUTNum_Reset = UUTNum_Reset - 1
        EraseThisUUT()
        mes = "UUT" & UUTNum & "has been erased by operator"
        UUTCount.Text = Str(UUTNum)
        StatusLog.Items.Add(mes)
        Me.Refresh()
    End Sub
    Private Sub EraseThisUUT()
        Dim SQLstr As String
        Dim Title As String
        Dim TraceID As Integer

        Title = Me.txtTitle.Text
        ExpectedProgress.Minimum = 0
        ActualProgress.Minimum = 0
        ActualProgress.Maximum = 4

        ILSetDone = False
        If MsgBox("Are you sure you want to erase UUT" & UUTNum_Reset, vbYesNo, "Cannot be undone.") = vbYes Then

            ResetLot()
            Me.txtTitle.Text = "     DELETING  UUT" & UUTNum_Reset & " TEST DATA"
            SQLstr = "Delete from TestData where JobNumber = '" & Me.cmbJob.Text & "' and SerialNumber = UUT" & UUTNum_Reset & "' and artwork_rev = '" & Me.txtArtwork.Text & "'"
            SQL.ExecuteSQLCommand(SQLstr, "NetworkData")
            If Not TweakMode Then If ActualProgress.Value + 1 <= ActualProgress.Maximum Then ActualProgress.Value = ExpectedProgress.Value + 1

            'Get Trace ID
            Me.txtTitle.Text = "     DELETING  UUT" & UUTNum_Reset & " Trace DATA"
            SQLstr = "select * from Trace where JobNumber = '" & Me.cmbJob.Text & "' and SerialNumber = UUT" & UUTNum_Reset & "' and artwork_rev = '" & Me.txtArtwork.Text & "'"
            TraceID = GetTestID(SQLstr, "NetworkTraceData")

            SQLstr = "Delete from TracePoints where TraceID = " & TraceID & ""
            SQL.ExecuteSQLCommand(SQLstr, "NetworkTraceData")
            If Not TweakMode Then If ActualProgress.Value + 1 <= ActualProgress.Maximum Then ActualProgress.Value = ExpectedProgress.Value + 1

            Me.txtTitle.Text = "     DELETING  UUT" & UUTNum_Reset & " TRACE DATA"
            SQLstr = "Delete from Trace where JobNumber = '" & Me.cmbJob.Text & "' and SerialNumber = UUT" & UUTNum_Reset & "' and artwork_rev = '" & Me.txtArtwork.Text & "'"
            If Not TweakMode Then If ActualProgress.Value + 1 <= ActualProgress.Maximum Then ActualProgress.Value = ExpectedProgress.Value + 1
        End If

        Me.txtTitle.Text = Title
    End Sub
    Private Sub EraseThisTest()
        Dim SQLstr As String
        Dim Title As String
        Dim TraceID As Integer

        Title = Me.txtTitle.Text
        ExpectedProgress.Minimum = 0
        ActualProgress.Minimum = 0
        ActualProgress.Maximum = 4

        ILSetDone = False
        SQLstr = "select * from Trace where JobNumber = '" & Me.cmbJob.Text & "'"
        TraceID = GetTestID(SQLstr, "LocalTraceData")

        If MsgBox("Are you sure you want to erase ALL DATA from This Job", vbYesNo, "Cannot be undone.") = vbYes Then

            ResetLot()
            Me.txtTitle.Text = "     DELETING " & Me.cmbJob.Text & " TEST DATA"
            SQLstr = "Delete from TestData where JobNumber = '" & Me.cmbJob.Text & "'"
            SQL.ExecuteSQLCommand(SQLstr, "NetworkData")
            If Not TweakMode Then If ActualProgress.Value + 1 <= ActualProgress.Maximum Then ActualProgress.Value = ExpectedProgress.Value + 1

            'Get Trace ID
            Me.txtTitle.Text = "     DELETING " & Me.cmbJob.Text & " Trace DATA"
            SQLstr = "select * from Trace where JobNumber = '" & Me.cmbJob.Text & "'"
            TraceID = GetTestID(SQLstr, "NetworkTraceData")

            SQLstr = "Delete from TracePoints where TraceID = " & TraceID & ""
            SQL.ExecuteSQLCommand(SQLstr, "NetworkTraceData")
            If Not TweakMode Then If ActualProgress.Value + 1 <= ActualProgress.Maximum Then ActualProgress.Value = ExpectedProgress.Value + 1

            Me.txtTitle.Text = "     DELETING " & Me.cmbJob.Text & " TRACE DATA"
            SQLstr = "Delete from Trace where JobNumber = '" & Me.cmbJob.Text & "'"
            If Not TweakMode Then If ActualProgress.Value + 1 <= ActualProgress.Maximum Then ActualProgress.Value = ExpectedProgress.Value + 1

            If MsgBox("Do you want to erase the Specification?", vbYesNo, "Cannot be undone.") = vbYes Then
                Me.txtTitle.Text = "     DELETING " & Me.cmbJob.Text & " Specifications"
                SQLstr = "Delete from Specifications where JobNumber = '" & Me.cmbJob.Text & "'"
                SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
                If Not TweakMode Then If ActualProgress.Value + 1 <= ActualProgress.Maximum Then ActualProgress.Value = ExpectedProgress.Value + 1

            End If
        End If

        Me.txtTitle.Text = Title
    End Sub

    Private Sub GetOffsets_Click(sender As Object, e As EventArgs) Handles GetOffsets.Click
        If GetOffsets.Checked And Not RunningOffsets Then
            cmdStartTest.Text = " Find  Offsets"
        Else
            cmdStartTest.Text = " Start   Test"
        End If

    End Sub
    Private Sub MasterUpLoad_Click(sender As Object, e As EventArgs) Handles GetOffsets.Click

    End Sub
    Private Sub MutiCal_Click(sender As Object, e As EventArgs) Handles MutiCal.Click
        If MutiCal.Checked Then
            Me.GetOffsets.Visible = False
            StatusLog.Items.Add("Multical Mode ON:" & DateTime.Now.ToString)
        Else
            Me.GetOffsets.Visible = True
            StatusLog.Items.Add("Multical Mode OFF:" & DateTime.Now.ToString)
        End If
    End Sub

    Private Sub AutotestToolStripSetupTrace_Click(sender As Object, e As EventArgs) Handles AutotestToolStripSetupTrace.Click
        Dim pos As Integer
        Dim SaveMessage As String
        SaveMessage = Me.txtTitle.Text
        Me.txtTitle.Text = "    PLEASE WAIT....... SETTING UP THE NETWORK ANALYZER...........  "
        System.Threading.Thread.Sleep(1)
        Me.Refresh()
        TEST1PASS = True
        TEST2PASS = True
        TEST3PASS = True
        TEST4PASS = True
        TEST5PASS = True
        If Me.MutiCal.Checked = False Then
            Tests.SetupVNA(True, 1)
        Else
            pos = cmbSwitch.SelectedIndex + 1
            Tests.SetupVNA(True, pos)
        End If
        ILSetDone = False
        Me.txtTitle.Text = SaveMessage
        Me.Refresh()
    End Sub
    Private Sub Simulation_Click(sender As Object, e As EventArgs) Handles Simulation.Click
        If Simulation.Checked Then
            Me.RealData.Visible = True
            Me.Pass.Visible = True
            Me.Fail.Visible = True
            Debug = True
            StatusLog.Items.Add("Simulation Mode ON:" & DateTime.Now.ToString)
        Else
            Me.RealData.Visible = False
            Me.Pass.Visible = False
            Me.Fail.Visible = False
            Debug = False
            StatusLog.Items.Add("Simulation Mode OFF:" & DateTime.Now.ToString)
        End If
    End Sub

    Private Sub EndLot_Click(sender As Object, e As EventArgs) Handles EndLot.Click
        Dim Temp As Double
        Dim mes As String
        StatusLog.Items.Add("Creating Report:" & DateTime.Now.ToString)
        ReportServer("report queue", UUTCount.Text)
        'ExcelData()
        txtCurrentTime.Text = "COMPLETE"
        ExpectedTimer.Stop()
        ClearStatusLog()
        StatusLog.Items.Add("Complete" & "   " & DateTime.Now.ToString)
        FirstPart = False
        Temp = Math.Round(ActualProgress.Value / ExpectedProgress.Value * 100, 0)
        If Temp >= 80 Then
            mes = "Congradulations: " & Temp & "% Efficient"
        ElseIf Temp >= 60 Then
            mes = "So so: " & Temp & "% Efficient"
        Else
            mes = "Not Good: " & Temp & "% Efficient"
        End If
        If ExpectedProgress2 > 0 Then SQL.UpdateEffeciency("Complete", txtEfficiency.Text, txtCurrentTime.Text, UUTCount.Text)
        StatusLog.Items.Add(mes)
        ResetLot()
        cmbJob.Text = " "
    End Sub
    Private Sub Form_Unload(Cancel As Integer)
        End
    End Sub
    Private Sub ClearFailureLog()
        If FailureLog.Items.Count > 0 Then
            For i% = 0 To FailureLog.Items.Count - 1
                FailureLog.Items.Remove(i)
            Next i%
        End If

        FailureLog.Refresh()
    End Sub
    Private Sub ClearStatusLog()
        If StatusLog.Items.Count > 0 Then
            For i% = 0 To StatusLog.Items.Count - 1
                StatusLog.Items.Remove(i)
            Next i%
        End If
        StatusLog.Refresh()
    End Sub
    Public Sub LoadJob(Job As String, Part As String)
        Me.cmbJob.Text = Job
        Me.cmbPart.Text = Part
    End Sub
    Private Sub UpdateFailureLog(Test1 As String, TEST2 As String, TEST3 As String, TEST4 As String, TEST5 As String)
        Dim Index As Integer
        If TweakMode Then Exit Sub
        If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then
            If Not TEST2PASS Or Not TEST1PASS Or Not TEST3PASS Or Not TEST4PASS Or Not TEST5PASS Then
                FailureLog.Items.Add("UUT Number: " & UUTNum & "         Job Number: " & Me.cmbJob.Text & "    Part Number: " & Me.cmbPart.Text & "              Insertion Loss: " & Test1 & "      Return Loss: " & TEST2 & "      Isolation: " & TEST3 & "      Amplitude Balance: " & TEST4 & "      Phase Balance: " & TEST5)
            End If
        ElseIf InStr(SpecType, "DIRECTIONAL COUPLER") Then
            If Not TEST2PASS Or Not TEST1PASS Or Not TEST3PASS Or Not TEST4PASS Or Not TEST5PASS Then
                FailureLog.Items.Add("UUT Number: " & UUTNum & "         Job Number: " & Me.cmbJob.Text & "    Part Number: " & Me.cmbPart.Text & "              Insertion Loss: " & Test1 & "      Return Loss: " & TEST2 & "      Coupling: " & TEST3 & "      Directivity: " & TEST4 & "      Coupled Balance: " & TEST5)
            End If
        End If
        Index = FailureLog.Items.Count
    End Sub

    Private Function RemoveLastEntryFailureLog() As Boolean
        Dim Index As Integer
        Index = FailureLog.Items.Count
        If Index <> 0 Then FailureLog.Items.RemoveAt(Index - 1)
        SQL.SaveFailureLog(" ")
        RemoveLastEntryFailureLog = True
        Me.Refresh()
    End Function

    Private Sub RemoveSerialEntryLog(UUT As String)
        Dim Index As Integer
        Dim Serial As String
        Serial = "UUT Number:" & UUT

        For Index = 0 To (FailureLog.TopIndex - 1)
            FailureLog.SelectedIndex = Index
            If InStr(FailureLog.Text, Serial) Then FailureLog.Items.RemoveAt(Index)
        Next Index
    End Sub

    Private Sub status(Color As String, Test As String, Optional Retest As Boolean = False)
        If Test = "TEST1" Then
            If Color = "Green" Then
                If UUTFail = 0 And Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.StartTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.ReTestFrame.BackColor = Drawing.Color.LawnGreen

                If UUTFail = 0 Or Retest Then Me.UUTLabel.ForeColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTCount.ForeColor = Drawing.Color.LawnGreen
                Me.PF1.ForeColor = Drawing.Color.LawnGreen
                If IL_TF Then
                    Me.Data1L.ForeColor = Drawing.Color.LawnGreen
                    Me.Data1H.ForeColor = Drawing.Color.LawnGreen
                Else
                    Me.Data1.ForeColor = Drawing.Color.LawnGreen
                End If
                If Not Retest Then Me.Failures1.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.FailTotal1.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                If Not Retest Then Me.Failures1.ForeColor = Drawing.Color.Red
                If Not Retest Then Me.FailTotal1.ForeColor = Drawing.Color.Red
                Me.PF1.ForeColor = Drawing.Color.Red
                If IL_TF Then
                    If IL1_status = "Pass" Then
                        Me.Data1L.ForeColor = Drawing.Color.LawnGreen
                    Else
                        Me.Data1L.ForeColor = Drawing.Color.Red
                    End If
                    If IL2_status = "Pass" Then
                        Me.Data1H.ForeColor = Drawing.Color.LawnGreen
                    Else
                        Me.Data1H.ForeColor = Drawing.Color.Red
                    End If
                Else
                    Me.Data1.ForeColor = Drawing.Color.Red
                End If
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
                Me.StartTestFrame.BackColor = Drawing.Color.Red
                Me.ReTestFrame.BackColor = Drawing.Color.Red
                Me.UUTLabel.ForeColor = Drawing.Color.Red
                Me.UUTCount.ForeColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF1.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF1.ForeColor = Drawing.Color.Black
                If IL_TF Then
                    Me.Data1L.ForeColor = Drawing.Color.Black
                    Me.Data1H.ForeColor = Drawing.Color.Black
                Else
                    Me.Data1.ForeColor = Drawing.Color.Black
                End If
                Me.ReTestFrame.BackColor = Drawing.Color.Black
                Me.UUTLabel.ForeColor = Drawing.Color.Black
                Me.StartTestFrame.BackColor = Drawing.Color.Black
                Me.PF1.ForeColor = Drawing.Color.Black
            End If

        End If

        If Test = "TEST2" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.StartTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.ReTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTLabel.ForeColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTCount.ForeColor = Drawing.Color.LawnGreen
                Me.PF2.ForeColor = Drawing.Color.LawnGreen
                Me.Data2.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.Failures2.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.FailTotal2.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                If Not Retest Then Me.Failures2.ForeColor = Drawing.Color.Red
                If Not Retest Then Me.FailTotal2.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
                Me.StartTestFrame.BackColor = Drawing.Color.Red
                Me.ReTestFrame.BackColor = Drawing.Color.Red
                Me.UUTLabel.ForeColor = Drawing.Color.Red
                Me.UUTCount.ForeColor = Drawing.Color.Red
                Me.PF2.ForeColor = Drawing.Color.Red
                Me.Data2.ForeColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF2.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF2.ForeColor = Drawing.Color.Black
                Me.Data2.ForeColor = Drawing.Color.Black
                Me.ReTestFrame.BackColor = Drawing.Color.Black
                Me.UUTLabel.ForeColor = Drawing.Color.Black
                Me.StartTestFrame.BackColor = Drawing.Color.Black
                Me.PF2.ForeColor = Drawing.Color.Black
            End If
        End If
        If Test = "TEST3" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.StartTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.ReTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTLabel.ForeColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTCount.ForeColor = Drawing.Color.LawnGreen
                Me.PF3.ForeColor = Drawing.Color.LawnGreen
                Me.Data3.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.Failures3.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.FailTotal3.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                If Not Retest Then Me.Failures3.ForeColor = Drawing.Color.Red
                If Not Retest Then Me.FailTotal3.ForeColor = Drawing.Color.Red
                Me.PF3.ForeColor = Drawing.Color.Red
                Me.Data3.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
                Me.StartTestFrame.BackColor = Drawing.Color.Red
                Me.ReTestFrame.BackColor = Drawing.Color.Red
                Me.UUTLabel.ForeColor = Drawing.Color.Red
                Me.UUTCount.ForeColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF3.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF3.ForeColor = Drawing.Color.Black
                Me.Data3.ForeColor = Drawing.Color.Black
                Me.ReTestFrame.BackColor = Drawing.Color.Black
                Me.UUTLabel.ForeColor = Drawing.Color.Black
                Me.UUTCount.ForeColor = Drawing.Color.Black
                Me.StartTestFrame.BackColor = Drawing.Color.Black
                Me.PF3.ForeColor = Drawing.Color.Black
            End If
        End If
        If Test = "TEST3L" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.StartTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.ReTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTLabel.ForeColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTCount.ForeColor = Drawing.Color.LawnGreen
                Me.PF3.ForeColor = Drawing.Color.LawnGreen
                Me.Data3L.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.Failures3.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.FailTotal3.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                If Not Retest Then Me.Failures3.ForeColor = Drawing.Color.Red
                If Not Retest Then Me.FailTotal3.ForeColor = Drawing.Color.Red
                Me.PF3.ForeColor = Drawing.Color.Red
                Me.Data3L.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
                Me.StartTestFrame.BackColor = Drawing.Color.Red
                Me.ReTestFrame.BackColor = Drawing.Color.Red
                Me.UUTLabel.ForeColor = Drawing.Color.Red
                Me.UUTCount.ForeColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF3.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF3.ForeColor = Drawing.Color.Black
                Me.Data3L.ForeColor = Drawing.Color.Black
                Me.ReTestFrame.BackColor = Drawing.Color.Black
                Me.UUTLabel.ForeColor = Drawing.Color.Black
                Me.StartTestFrame.BackColor = Drawing.Color.Black
                Me.PF3.ForeColor = Drawing.Color.Black
            End If
        End If
        If Test = "TEST3H" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.StartTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.ReTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTLabel.ForeColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTCount.ForeColor = Drawing.Color.LawnGreen
                Me.PF3.ForeColor = Drawing.Color.LawnGreen
                Me.Data3H.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.Failures3.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.FailTotal3.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                If Not Retest Then Me.Failures3.ForeColor = Drawing.Color.Red
                If Not Retest Then Me.FailTotal3.ForeColor = Drawing.Color.Red
                Me.PF3.ForeColor = Drawing.Color.Red
                Me.Data3H.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
                Me.StartTestFrame.BackColor = Drawing.Color.Red
                Me.ReTestFrame.BackColor = Drawing.Color.Red
                Me.UUTLabel.ForeColor = Drawing.Color.Red
                Me.UUTCount.ForeColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF3.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF3.ForeColor = Drawing.Color.Black
                Me.Data4H.ForeColor = Drawing.Color.Black
                Me.ReTestFrame.BackColor = Drawing.Color.Black
                Me.UUTLabel.ForeColor = Drawing.Color.Black
                Me.StartTestFrame.BackColor = Drawing.Color.Black
                Me.PF3.ForeColor = Drawing.Color.Black
            End If
        End If

        If Test = "TEST4" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.StartTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.ReTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTLabel.ForeColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTCount.ForeColor = Drawing.Color.LawnGreen
                Me.PF4.ForeColor = Drawing.Color.LawnGreen
                Me.Data4.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.Failures4.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.FailTotal4.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                If Not Retest Then Me.Failures4.ForeColor = Drawing.Color.Red
                If Not Retest Then Me.FailTotal4.ForeColor = Drawing.Color.Red
                Me.PF4.ForeColor = Drawing.Color.Red
                Me.Data4.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
                Me.StartTestFrame.BackColor = Drawing.Color.Red
                Me.ReTestFrame.BackColor = Drawing.Color.Red
                Me.UUTLabel.ForeColor = Drawing.Color.Red
                Me.UUTCount.ForeColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF4.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF4.ForeColor = Drawing.Color.Black
                Me.Data4.ForeColor = Drawing.Color.Black
                Me.ReTestFrame.BackColor = Drawing.Color.Black
                Me.UUTLabel.ForeColor = Drawing.Color.Black
                Me.StartTestFrame.BackColor = Drawing.Color.Black
                Me.PF4.ForeColor = Drawing.Color.Black
            End If
        End If
        If Test = "TEST4L" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.StartTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.ReTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTLabel.ForeColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTCount.ForeColor = Drawing.Color.LawnGreen
                Me.PF4.ForeColor = Drawing.Color.LawnGreen
                Me.Data4L.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.Failures4.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.FailTotal4.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                If Not Retest Then Me.Failures4.ForeColor = Drawing.Color.Red
                If Not Retest Then Me.FailTotal4.ForeColor = Drawing.Color.Red
                Me.PF4.ForeColor = Drawing.Color.Red
                Me.Data4L.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
                Me.StartTestFrame.BackColor = Drawing.Color.Red
                Me.ReTestFrame.BackColor = Drawing.Color.Red
                Me.UUTLabel.ForeColor = Drawing.Color.Red
                Me.UUTCount.ForeColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF4.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF4.ForeColor = Drawing.Color.Black
                Me.Data4L.ForeColor = Drawing.Color.Black
                Me.ReTestFrame.BackColor = Drawing.Color.Black
                Me.UUTLabel.ForeColor = Drawing.Color.Black
                Me.StartTestFrame.BackColor = Drawing.Color.Black
                Me.PF4.ForeColor = Drawing.Color.Black
            End If
        End If
        If Test = "TEST4H" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.StartTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.ReTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTLabel.ForeColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTCount.ForeColor = Drawing.Color.LawnGreen
                Me.PF4.ForeColor = Drawing.Color.LawnGreen
                Me.Data4H.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.Failures4.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.FailTotal4.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                If Not Retest Then Me.Failures4.ForeColor = Drawing.Color.Red
                If Not Retest Then Me.FailTotal4.ForeColor = Drawing.Color.Red
                Me.PF4.ForeColor = Drawing.Color.Red
                Me.Data4H.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
                Me.StartTestFrame.BackColor = Drawing.Color.Red
                Me.ReTestFrame.BackColor = Drawing.Color.Red
                Me.UUTLabel.ForeColor = Drawing.Color.Red
                Me.UUTCount.ForeColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF4.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF4.ForeColor = Drawing.Color.Black
                Me.Data4H.ForeColor = Drawing.Color.Black
                Me.ReTestFrame.BackColor = Drawing.Color.Black
                Me.UUTLabel.ForeColor = Drawing.Color.Black
                Me.StartTestFrame.BackColor = Drawing.Color.Black
                Me.PF4.ForeColor = Drawing.Color.Black
            End If
        End If

        If Test = "TEST5" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.StartTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.ReTestFrame.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTLabel.ForeColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Or Retest Then Me.UUTCount.ForeColor = Drawing.Color.LawnGreen
                Me.PF5.ForeColor = Drawing.Color.LawnGreen
                Me.Data5.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.Failures5.ForeColor = Drawing.Color.LawnGreen
                If Not Retest Then Me.FailTotal5.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                If Not Retest Then Me.Failures5.ForeColor = Drawing.Color.Red
                If Not Retest Then Me.FailTotal5.ForeColor = Drawing.Color.Red
                Me.PF5.ForeColor = Drawing.Color.Red
                Me.Data5.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
                Me.StartTestFrame.BackColor = Drawing.Color.Red
                Me.ReTestFrame.BackColor = Drawing.Color.Red
                Me.UUTLabel.ForeColor = Drawing.Color.Red
                Me.UUTCount.ForeColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF5.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF5.ForeColor = Drawing.Color.Black
                Me.Data5.ForeColor = Drawing.Color.Black
                Me.ReTestFrame.BackColor = Drawing.Color.Black
                Me.UUTLabel.ForeColor = Drawing.Color.Black
                Me.StartTestFrame.BackColor = Drawing.Color.Black
                Me.PF5.ForeColor = Drawing.Color.Black
            End If
        End If
    End Sub

    '  Private Sub Random_Click()
    '     If Random.Checked Then
    '          Me.Pass.Checked = False
    ''          Me.Fail.Checked = False
    '     End If
    ' End Sub

    Private Sub Pass_Click(sender As Object, e As EventArgs) Handles Pass.Click
        If Pass.Checked Then
            Me.RealData.Checked = False
            Me.Fail.Checked = False
        End If
    End Sub

    Private Sub Fail_Click(sender As Object, e As EventArgs) Handles Fail.Click
        If Fail.Checked Then
            Me.RealData.Checked = False
            Me.Pass.Checked = False
        End If
    End Sub

    Private Sub TraceSearch_Click()
        'Dim Temp
        ' frmTraceSearch.Left = -60
        ' frmTraceSearch.Top = 2740
        ' frmTraceSearch.Width = 25320
        ' frmTraceSearch.Height = 12320

        'frmTraceSearch.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
        'frmTraceSearch.Top = GetSetting(App.Title, "Settings", "MainTop", 1000) + 2800
        'frmTraceSearch.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
        'frmTraceSearch.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500) - 3100
        ' frmTraceSearch.Show(vbModeless, Me)
    End Sub

    Private Function ResumeTesting() As Boolean
        Dim SQLstr As String
        Dim TestExist As Boolean
        Dim RetrnStr As String
        Dim OldTitle As String
        Dim TempAccess As Boolean
        Dim TempNetwork As Boolean
        Dim PassFail As String
        Dim thisStatus1 As String = ""
        Dim thisStatus2 As String = ""
        Dim thisStatus3 As String = ""
        Dim thisStatus As String = ""
        Dim thisUser As String = ""
        Dim thisJob As String = ""
        Dim thisStation As String = ""
        Dim complete As String = ""

        Try
            TestExist = False
            OldTitle = txtTitle.Text
            ResetLot()

            UUTNum = 0
            ResumeTesting = False
            If Me.cmbJob.Text = " " Then Exit Function
            TempAccess = SQLAccess
            TempNetwork = NetworkAccess
            Dim test_status As String = GetTestStatus()
            stopTest = False
            SQLstr = "select * from TestData where JobNumber = '" & Me.cmbJob.Text & "' And WorkStation = '" & WorkStation & "'"
            If test_status = "None" Then
                If SQL.CheckforRow(SQLstr, "NetworkData") > 0 Then GoTo StartResume
            ElseIf test_status = "One" Then
                If ReportStatus(0) = "test running" And WorkStation = SavedWorkStation(0) And SavedJob(0) = Job Then
                    SQLstr = "select * from TestData where JobNumber = '" & Me.cmbJob.Text & "'"
                    If SQL.CheckforRow(SQLstr, "NetworkData") > 0 Then GoTo StartResume
                ElseIf ReportStatus(0) = "test running" And WorkStation = SavedWorkStation(0) And Not SavedJob(0) = Job Then
                    If MsgBox("Are you sure you  want to continue at UUT" & SavedComplete(0) & "??", vbYesNo, SavedJob(0) & " has not been completed by " & SavedUser(0) & " at Workstation " & SavedWorkStation(0)) = vbYes Then
                        If MsgBox("Do you want to close " & SavedJob(0) & "from Workstation " & SavedWorkStation(0) & "?", vbYesNo, "Yes or No") = vbYes Then
                            SQLstr = "UPDATE ReportQueue Set ReportStatus = job closed where JobNumber = '" & Job & "'"
                            SQL.ExecuteSQLCommand(SQLstr, "Effeciency")
                            SQLstr = "select * from TestData where JobNumber = '" & Me.cmbJob.Text & "''"
                            If SQL.CheckforRow(SQLstr, "NetworkData") > 0 Then GoTo StartResume
                            stopTest = False
                        Else
                            stopTest = True
                        End If
                    End If
                ElseIf ReportStatus(0) = "report complete" Or ReportStatus(0) = "report complete" Then
                    If MsgBox("Are you sure you  want to continue at UUT" & SavedComplete(0) & "??", vbYesNo, SavedJob(0) & " has been completed by " & SavedUser(0) & " at Workstation " & SavedWorkStation(0)) = vbYes Then
                        SQLstr = "select * from TestData where JobNumber = '" & Me.cmbJob.Text & "'"
                        If SQL.CheckforRow(SQLstr, "NetworkData") > 0 Then GoTo StartResume
                        GoTo StartResume
                    End If
                ElseIf ReportStatus(0) = "job closed" Then
                    If MsgBox("Are you sure you  want to continue at UUT" & SavedComplete(0) & "??", vbYesNo, SavedJob(0) & " has been paused by " & SavedUser(0) & " at Workstation " & SavedWorkStation(0)) = vbYes Then
                        GoTo StartResume
                    Else
                        stopTest = True
                    End If
                Else
                    GoTo StartResume
                End If
            ElseIf test_status = "Multiple" Then
                complete = 0
                For x = 0 To ReportStatus.Length - 1
                    thisStatus1 = ReportStatus(x)
                    If thisStatus1 = "job closed" Then
                        thisStatus = ReportStatus(x)
                        If SavedComplete(x) > complete Then
                            complete = SavedComplete(x)
                        End If
                        thisUser = SavedUser(x)
                        thisStation = SavedWorkStation(x)
                        thisJob = SavedJob(x)
                        Exit For
                    Else
                        thisStatus = thisStatus1
                    End If

                Next
                If Not complete = 0 Then GoTo letsgetiton
                complete = 0
                For x = 0 To ReportStatus.Length - 1
                    thisStatus2 = ReportStatus(x)
                    If thisStatus2 = "test running" Then
                        thisStatus = ReportStatus(x)
                        If SavedComplete(x) > complete Then
                            complete = SavedComplete(x)
                        End If
                        thisUser = SavedUser(x)
                        thisStation = SavedWorkStation(x)
                        thisJob = SavedJob(x)
                        Exit For
                    End If
                Next
                If Not complete = 0 Then GoTo letsgetiton
                complete = 0
                For x = 0 To ReportStatus.Length - 1
                    thisStatus2 = ReportStatus(x)
                    If thisStatus2 = "report queue" Then
                        thisStatus = ReportStatus(x)
                        If SavedComplete(x) > complete Then
                            complete = SavedComplete(x)
                        End If
                        thisUser = SavedUser(x)
                        thisStation = SavedWorkStation(x)
                        thisJob = SavedJob(x)
                        Exit For
                    End If
                Next
                If Not complete = 0 Then GoTo letsgetiton
                complete = 0
                For x = 0 To ReportStatus.Length - 1
                    thisStatus3 = ReportStatus(x)
                    If thisStatus3 = "job closed" Then
                        thisStatus = ReportStatus(x)
                        If SavedComplete(x) > complete Then
                            complete = SavedComplete(x)
                        End If
                        thisUser = SavedUser(x)
                        thisStation = SavedWorkStation(x)
                        thisJob = SavedJob(x)
                        Exit For
                    End If
                Next
letsgetiton:
                If MsgBox("Are you sure you  want to continue at UUT" & complete & " ??", vbYesNo, thisJob & " has not been completed by multiple Users and/or Workstations") = vbYes Then
                    SQLstr = "UPDATE ReportQueue Set ReportStatus = job closed where JobNumber = '" & Job & "'"
                    SQL.ExecuteSQLCommand(SQLstr, "Effeciency")
                    SQLstr = "select * from TestData where JobNumber = '" & Me.cmbJob.Text & "'"
                    If SQL.CheckforRow(SQLstr, "NetworkData") > 0 Then GoTo StartResume
                    stopTest = False
                Else
                    If MsgBox("Do you want to close " & SavedJob(0) & "?", vbYesNo, "Yes or No") = vbYes Then

                        If MsgBox("Do you want to erase the data from This Job", vbYesNo, "Cannot be undone.") = vbYes Then
                            EraseThisTest()
                        End If
                        SQLstr = "select * from TestData where JobNumber = '" & Me.cmbJob.Text & "'"
                        If SQL.CheckforRow(SQLstr, "NetworkData") > 0 Then GoTo StartResume
                        stopTest = False
                    Else
                        stopTest = True
                    End If
                End If
            End If
            SQLAccess = TempAccess
            NetworkAccess = TempNetwork
            GoTo NoResume

StartResume:
            SQLAccess = TempAccess
            NetworkAccess = TempNetwork
            ResumeTesting = True
            stopTest = False
            TestExist = True
            If TweakMode Then Exit Function
            If MsgBox("The Job is already on record. Resume Testing", vbYesNo, "Continue or Start Over") = vbNo Then
                If MsgBox("Do you want to erase The data?", vbYesNo, "Erasing Job Data") = vbNo Then
                    Exit Function
                Else
                    EraseThisTest()
                    Exit Function
                End If
            Else
                ResumeTesting = True
                txtTitle.Text = " Updating Test Run Please Wait "
                If SpecAB_TF Then
                    Data4L.Visible = True
                    Data4H.Visible = True
                    Data4.Visible = False
                Else
                    Data4L.Visible = False
                    Data4H.Visible = False
                    Data4.Visible = True
                End If
                If SQLAccess Then
                    Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                    Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                    ats.Open()
                    System.Threading.Thread.Sleep(20)
                    Me.Refresh()
                    Dim dr As SqlDataReader = cmd.ExecuteReader()
                    While Not dr.Read = Nothing
                        Me.RunStatus.ForeColor = Color.Red
                        UUTNum = UUTNum + 1
                        UUTNum_Reset = UUTNum_Reset + 1
                        EraseTest.Text = "Remove UUT" & UUTNum - 1
                        If UUTNum > 5 And Not BypassUnchecked Then Me.GetTrace.Checked = False
                        UUTCount.Text = Str(UUTNum)
                        If TweakMode Then
                            Me.UUTMessage.Text = "  UUT TESTS  --   Testing Undisclosed Unit"
                        ElseIf Not TraceChecked Then
                            UUTMessage.Text = "  UUT TESTS Marker Mode  --   Load Unit #" & UUTNum + 1
                        Else
                            Me.UUTMessage.Text = "  UUT TESTS  --   Testing Unit #" & UUTNum
                        End If
                        Me.cmdStartTest.Text = "UUT" & UUTNum
                        Me.cmdStartTest.Enabled = False
                        ResetTests()

                        'Insertion Loss
                        If Not IsDBNull(dr.Item(6)) Then
                            RetrnVal = CDbl(dr.Item(6))
                        Else
                            If Not Incomplete Then
                                UUTNum = UUTNum - 1
                                UUTNum_Reset = UUTNum_Reset - 1
                                Me.txtTitle.Text = "     DELETED BAD DATA. UUT" & UUTNum_Reset
                                SQLstr = "Delete from TestData where JobNumber = '" & Me.cmbJob.Text & "' and SerialNumber = UUT" & UUTNum_Reset & "'"
                                SQL.ExecuteSQLCommand(SQLstr, "NetworkData")
                                Incomplete = True
                                GoTo Recap
                            End If
                        End If
                        PassFail = Tests.InsertionLoss3dB(True)
                        status("Blue", "TEST1")
                        PF1.Text = PassFail
                        RetrnStr = CStr(TruncateDecimal(RetrnVal, 2))
                        Data1.Text = Format(RetrnVal, "0.00")
                        If PassFail = "Pass" Then
                            TEST1PASS = True
                            status("Green", "TEST1")
                            MSChart.UpDateChartData(SpecType, "IL", "Pass")
                        ElseIf PassFail = "Fail" Then
                            TEST1PASS = False
                            status("Red", "TEST1")

                            MSChart.UpDateChartData(SpecType, "IL", "Fail")
                            UUTFail = 1
                        End If
                        Me.Failures1.Text = FormatPercent(((Test1Fail / UUTNum)), 1)
                        Me.Total1.Text = UUTNum
                        Me.FailTotal1.Text = Test1Fail

                        'Return Loss
                        If Not IsDBNull(dr.Item(7)) Then
                            RetrnVal = CDbl(dr.Item(7))
                        Else
                            If Not TEST2PASS Then
                                Test2Fail = Test2Fail - 1
                                status("Green", "TEST1")
                                MSChart.UpDateChartData(SpecType, "RL", "Pass")
                            End If
                            If Not TEST2PASS Then
                                If Test2Fail = 0 Then status("Green", "TEST2")
                                MSChart.UpDateChartData(SpecType, "RL", "Pass")
                            End If
                            If Not Incomplete Then
                                UUTNum = UUTNum - 1
                                UUTNum_Reset = UUTNum_Reset - 1
                                Me.txtTitle.Text = "     DELETED BAD DATA. UUT" & UUTNum_Reset
                                SQLstr = "Delete from TestData where JobNumber = '" & Me.cmbJob.Text & "' and SerialNumber = UUT" & UUTNum_Reset & "'"
                                SQL.ExecuteSQLCommand(SQLstr, "NetworkData")
                                Incomplete = True
                                GoTo Recap
                            End If
                        End If
                        PassFail = Tests.ReturnLoss(True)
                        RetrnStr = CStr(TruncateDecimal(RetrnVal, 2))
                        status("Blue", "TEST2")
                        PF2.Text = PassFail
                        RetrnStr = CStr(TruncateDecimal(RetrnVal, 2))
                        Data2.Text = Format(RetrnVal, "0.0")
                        If PassFail = "Pass" Then
                            TEST2PASS = True
                            status("Green", "TEST2")
                        ElseIf PassFail = "Fail" Then
                            TEST2PASS = False
                            status("Red", "TEST2")
                            UUTFail = 1
                        End If
                        Me.Failures2.Text = FormatPercent(((Test2Fail / UUTNum)), 1)
                        Me.Total2.Text = UUTNum
                        Me.FailTotal2.Text = Test2Fail

                        'Isolation
                        'Coupling
                        If SpecType.Contains("COMBINER/DIVIDER") And ISO_TF Then
                            ISoL = CDbl(dr.Item(9))
                            ISoH = CDbl(dr.Item(18))
                            If ISoL > ISoH Then
                                RetrnVal = ISoL
                            Else
                                RetrnVal = ISoH
                            End If
                        ElseIf (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And Not IsDBNull(dr.Item(9)) Then
                            RetrnVal = CDbl(dr.Item(9))
                        ElseIf (SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER") And Not IsDBNull(dr.Item(8)) Then
                            RetrnVal = CDbl(dr.Item(8))
                        Else
                            If Not TEST1PASS Then
                                If Test1Fail = 0 Then status("Green", "TEST1")
                                MSChart.UpDateChartData(SpecType, "IL", "Pass")
                            End If
                            If Not TEST2PASS Then
                                status("Green", "TEST2")
                                MSChart.UpDateChartData(SpecType, "RL", "Pass")
                            End If
                            If Not TEST3PASS Then
                                status("Green", "TEST3")
                                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then MSChart.UpDateChartData(SpecType, "ISO", "Pass")
                                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then MSChart.UpDateChartData(SpecType, "COU", "Pass")
                            End If
                            If Not Incomplete Then
                                UUTNum = UUTNum - 1
                                UUTNum_Reset = UUTNum_Reset - 1
                                Me.txtTitle.Text = "     DELETED BAD DATA. UUT" & UUTNum_Reset
                                SQLstr = "Delete from TestData where JobNumber = '" & Me.cmbJob.Text & "' and SerialNumber = UUT" & UUTNum_Reset & "'"
                                SQL.ExecuteSQLCommand(SQLstr, "NetworkData")
                                Incomplete = True
                                GoTo Recap
                            End If
                        End If
                        PassFail = Tests.Isolation(RetrnVal, True)
                        RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                        If ISO_TF Then
                            status("Blue", "TEST3L")
                            status("Blue", "TEST3H")
                        Else
                            status("Blue", "TEST3")
                        End If
                        PF3.Text = PassFail
                        Data3.Text = Format(RetrnVal, "0.0")
                        If PassFail = "Pass" Then
                            TEST3PASS = True
                            status("Green", "TEST3")
                        ElseIf PassFail = "Fail" Then
                            TEST3PASS = False
                            status("Red", "TEST3")
                            UUTFail = 1
                        End If
                        Me.Failures3.Text = FormatPercent(((TEST3Fail / UUTNum)), 1)
                        Me.Total3.Text = UUTNum
                        Me.FailTotal3.Text = TEST3Fail

                        'AmplitudeBalance
                        'Directivity
                        If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And Not IsDBNull(dr.Item(11)) Then
                            If SpecAB_TF Then
                                AB1 = CDbl(dr.Item(16))
                                AB2 = CDbl(dr.Item(17))
                            Else
                                RetrnVal = CDbl(dr.Item(11))
                            End If

                        ElseIf (SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER") And Not IsDBNull(dr.Item(10)) Then
                            RetrnVal = CDbl(dr.Item(10))
                        Else
                            If Not TEST1PASS Then
                                If Test1Fail = 0 Then status("Green", "TEST1")
                                ' MSChart.UpDateChartData(SpecType, "IL", "Pass")
                            End If
                            If Not TEST2PASS Then
                                status("Green", "TEST2")
                                'MSChart.UpDateChartData(SpecType, "RL", "Pass")
                            End If
                            If Not TEST3PASS Then
                                status("Green", "TEST3")
                                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then MSChart.UpDateChartData(SpecType, "ISO", "Pass")
                                If (SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER") Then MSChart.UpDateChartData(SpecType, "COU", "Pass")
                            End If
                            If Not TEST4PASS Then
                                status("Green", "TEST4")
                                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then MSChart.UpDateChartData(SpecType, "AB", "Pass")
                                If (SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER") Then MSChart.UpDateChartData(SpecType, "DIR", "Pass")
                            End If
                            If Not Incomplete Then
                                If MsgBox("Data is incomplete at UUT" & UUTNum, vbOKOnly, "Resume Testing At " & UUTNum) = vbOK Then
                                    UUTNum = UUTNum - 1
                                    EraseTest.Text = "Remove UUT" & UUTNum - 1
                                    UUTNum_Reset = UUTNum_Reset - 1
                                    Incomplete = True
                                    GoTo Recap
                                End If
                            End If
                        End If
                        If Not SpecType = "TRANSFORMER" Then
                            PassFail = Tests.AmplitudeBalance(True)
                            RetrnVal = AB
                            If SpecAB_TF Then
                                status("Blue", "TEST4L")
                                status("Blue", "TEST4H")
                            Else
                                status("Blue", "TEST4")
                            End If
                            PF4.Text = PassFail
                            If SpecAB_TF And Not AB1 = 0.0 Then
                                AB1 = CStr(TruncateDecimal(AB1, 2))
                                AB2 = CStr(TruncateDecimal(AB2, 2))
                                Data4L.Text = AB1
                                Data4H.Text = AB2
                                Me.Total4.Text = UUTNum
                                If AB1Pass = "Pass" Then
                                    status("Green", "TEST4L", True)
                                ElseIf AB1Pass = "Fail" Then
                                    status("Red", "TEST4L", True)
                                End If
                                If AB2Pass = "Pass" Then
                                    status("Green", "TEST4H", True)
                                ElseIf AB2Pass = "Fail" Then
                                    status("Red", "TEST4H", True)
                                End If
                            Else
                                RetrnStr = CStr(TruncateDecimal(RetrnVal, 2))
                                Data4.Text = Format(RetrnVal, "0.00")
                                If PassFail = "Pass" Then
                                    TEST4PASS = True
                                    status("Green", "TEST4")
                                ElseIf PassFail = "Fail" Then
                                    TEST4PASS = False

                                    status("Red", "TEST4")
                                    UUTFail = 1
                                End If
                            End If
                        End If

                        Me.Failures4.Text = FormatPercent(((TEST4Fail / UUTNum)), 1)
                        Me.Total4.Text = UUTNum
                        Me.FailTotal4.Text = TEST4Fail

                        'PhaseBalance
                        'Coupled Flatness
                        If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And Not IsDBNull(dr.Item(13)) Then
                            RetrnVal = CDbl(dr.Item(13))
                        ElseIf (SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER") And Not IsDBNull(dr.Item(12)) Then
                            RetrnVal = CDbl(dr.Item(12))
                        Else
                            If Not TEST1PASS Then
                                If Test1Fail = 0 Then status("Green", "TEST1")
                                MSChart.UpDateChartData(SpecType, "IL", "Pass")
                            End If
                            If Not TEST2PASS Then
                                status("Green", "TEST2")
                                MSChart.UpDateChartData(SpecType, "RL", "Pass")
                            End If
                            If Not TEST3PASS Then
                                status("Green", "TEST3")
                                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then MSChart.UpDateChartData(SpecType, "ISO", "Pass")
                                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then MSChart.UpDateChartData(SpecType, "COU", "Pass")
                            End If
                            If Not TEST4PASS Then
                                status("Green", "TEST4")
                                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then MSChart.UpDateChartData(SpecType, "AB", "Pass")
                                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then MSChart.UpDateChartData(SpecType, "DIR", "Pass")
                            End If
                            If Not TEST5PASS Then
                                status("Green", "TEST5")
                                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then MSChart.UpDateChartData(SpecType, "PB", "Pass")
                                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then MSChart.UpDateChartData(SpecType, "CB", "Pass")
                            End If
                            If Not Incomplete Then
                                If MsgBox("Data is incomplete at UUT" & UUTNum, vbOKOnly, "Resume Testing At " & UUTNum) = vbOK Then
                                    UUTNum = UUTNum - 1
                                    EraseTest.Text = "Remove UUT" & UUTNum - 1
                                    UUTNum_Reset = UUTNum_Reset - 1
                                    Incomplete = True
                                    GoTo Recap
                                End If
                            End If
                        End If
                        PassFail = Tests.PhaseBalance(SpecType, True)
                        status("Blue", "TEST5")
                        PF5.Text = PassFail
                        RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                        Data5.Text = Format(RetrnVal, "0.0")
                        If PassFail = "Pass" Then
                            TEST5PASS = True
                            status("Green", "TEST5")
                        ElseIf PassFail = "Fail" Then
                            TEST5PASS = False
                            status("Red", "TEST5")
                        End If
                        Me.Failures5.Text = FormatPercent(((TEST5Fail / UUTNum)), 1)
                        Me.Total5.Text = UUTNum
                        Me.FailTotal5.Text = TEST5Fail

                        If Not TEST2PASS Or Not TEST1PASS Or Not TEST3PASS Or Not TEST4PASS Or Not TEST5PASS Then
                            UpdateFailureLog(Me.PF1.Text, Me.PF2.Text, Me.PF3.Text, Me.PF4.Text, Me.PF5.Text)
                            LOTFail = LOTFail + 1
                            Me.Refresh()

                        End If
Recap:
                        If LOTFail = 0 Then LotTestFrame.BackColor = Color.LawnGreen
                        If LOTFail = 0 Then LotFailureFrame.BackColor = Color.LawnGreen
                        If LOTFail > 0 Then LotTestFrame.BackColor = Color.Red
                        If LOTFail > 0 Then LotFailureFrame.BackColor = Color.Red
                        If LOTFail > 0 Then UUTStatusColor.BackColor = Color.Red
                        If UUTFail = 1 Then Me.cmdRetest.Enabled = True
                        If UUTFail = 0 Then Me.cmdRetest.Enabled = False
                        If UUTFail = 1 Then Me.EraseTest.Enabled = True
                        If UUTFail = 0 Then Me.EraseTest.Enabled = False
                        If UUTNum <> 0 Then Me.LotFailure.Text = FormatPercent(((LOTFail / UUTNum)), 1)
                        If TweakMode Then
                            UUTMessage.Text = "  UUT TESTS  --   Tweak Mode. No Data Logging"
                        ElseIf Not TraceChecked Then
                            UUTMessage.Text = "  UUT TESTS Marker Mode  --   Load Unit #" & UUTNum + 1
                        Else
                            UUTMessage.Text = "  UUT TESTS  --   Load Unit #" & UUTNum + 1
                        End If

                        Me.cmdStartTest.Text = "Next UUT"
                        cmdStartTest.Enabled = True
                        ' Me.cmdRetest.Enabled = True

                        TEST1PASS = True
                        TEST2PASS = True
                        TEST3PASS = True
                        TEST4PASS = True
                        TEST5PASS = True
                        Me.Refresh()
                    End While
                Else
                    Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkData")
                    Dim atsLocal As New OleDb.OleDbConnection
                    atsLocal.ConnectionString = strConnectionString
                    Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                    atsLocal.Open()
                    System.Threading.Thread.Sleep(20)
                    Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader

                    While Not drLocal.Read = Nothing
                        Me.RunStatus.ForeColor = Color.Red
                        UUTNum = UUTNum + 1
                        EraseTest.Text = "Remove UUT" & UUTNum
                        UUTNum_Reset = UUTNum_Reset + 1
                        If UUTNum > 5 And Not BypassUnchecked Then Me.GetTrace.Checked = False
                        UUTCount.Text = Str(UUTNum)
                        If TweakMode Then
                            Me.UUTMessage.Text = "  UUT TESTS  --   Testing Undisclosed Unit"
                        ElseIf TraceChecked Then
                            UUTMessage.Text = "  UUT TESTS Marker Mode  --   Load Unit #" & UUTNum + 1
                        Else
                            Me.UUTMessage.Text = "  UUT TESTS  --   Testing Unit #" & UUTNum
                        End If
                        Me.cmdStartTest.Text = "UUT" & UUTNum
                        Me.cmdStartTest.Enabled = False
                        ResetTests()

                        'Insertion Loss
                        If drLocal.Item(6) IsNot Nothing Then
                            RetrnVal = CDbl(drLocal.Item(6))
                        Else
                            If Not Incomplete Then
                                If MsgBox("Data is incomplete at UUT" & UUTNum, vbOKOnly, "Resume Testing At " & UUTNum) = vbOK Then
                                    UUTNum = UUTNum - 1
                                    UUTNum_Reset = UUTNum_Reset - 1
                                    Incomplete = True
                                    GoTo Recap
                                End If
                            End If
                        End If
                        PassFail = Tests.InsertionLoss3dB(True)
                        status("Blue", "TEST1")
                        PF1.Text = PassFail
                        RetrnStr = CStr(TruncateDecimal(RetrnVal, 2))
                        Data1.Text = Format(RetrnVal, "0.00")
                        If PassFail = "Pass" Then
                            TEST1PASS = True
                            status("Green", "TEST1")
                        ElseIf PassFail = "Fail" Then
                            TEST1PASS = False
                            status("Red", "TEST1")
                            UUTFail = 1
                        End If
                        Me.Failures1.Text = FormatPercent(((Test1Fail / UUTNum)), 1)
                        Me.Total1.Text = UUTNum
                        Me.FailTotal1.Text = Test1Fail

                        'Return Loss
                        If Not IsDBNull(drLocal.Item(7)) Then
                            RetrnVal = CDbl(drLocal.Item(7))
                        Else
                            If Not TEST2PASS Then
                                status("Green", "TEST1")
                            End If
                            If Not TEST2PASS Then
                                status("Green", "TEST2")
                            End If
                            If Not Incomplete Then
                                If MsgBox("Data is incomplete at UUT" & UUTNum, vbOKOnly, "Resume Testing At " & UUTNum) = vbOK Then
                                    UUTNum = UUTNum - 1
                                    EraseTest.Text = "Remove UUT" & UUTNum
                                    UUTNum_Reset = UUTNum_Reset - 1
                                    Incomplete = True
                                    GoTo Recap
                                End If
                            End If
                        End If
                        PassFail = Tests.ReturnLoss(True)
                        status("Blue", "TEST2")
                        PF2.Text = PassFail
                        RetrnStr = CStr(TruncateDecimal(RetrnVal, 2))
                        Data2.Text = Format(RetrnVal, "0.0")
                        If PassFail = "Pass" Then
                            TEST2PASS = True
                            status("Green", "TEST2")
                        ElseIf PassFail = "Fail" Then
                            TEST2PASS = False
                            status("Red", "TEST2")
                            UUTFail = 1
                        End If
                        Me.Failures2.Text = FormatPercent(((Test2Fail / UUTNum)), 1)
                        Me.Total2.Text = UUTNum
                        Me.FailTotal2.Text = Test2Fail

                        'Isolation
                        'Coupling
                        If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And Not IsDBNull(drLocal.Item(9)) Then
                            RetrnVal = CDbl(drLocal.Item(9))
                        ElseIf (SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER") And Not IsDBNull(drLocal.Item(8)) Then
                            RetrnVal = CDbl(drLocal.Item(8))
                        Else
                            If Not TEST1PASS Then
                                If Test1Fail = 0 Then status("Green", "TEST1")
                            End If
                            If Not TEST2PASS Then
                                status("Green", "TEST2")
                            End If
                            If Not TEST3PASS Then
                                status("Green", "TEST3")
                            End If
                            If Not Incomplete Then
                                If MsgBox("Data is incomplete at UUT" & UUTNum, vbOKOnly, "Resume Testing At " & UUTNum) = vbOK Then
                                    UUTNum = UUTNum - 1
                                    EraseTest.Text = "Remove UUT" & UUTNum
                                    UUTNum_Reset = UUTNum_Reset - 1
                                    Incomplete = True
                                    GoTo Recap
                                End If
                            End If
                        End If
                        PassFail = Tests.Isolation(RetrnVal, True)
                        status("Blue", "TEST3")
                        PF3.Text = PassFail
                        RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                        Data3.Text = Format(RetrnVal, "0.0")
                        If PassFail = "Pass" Then
                            TEST3PASS = True
                            status("Green", "TEST3")
                        ElseIf PassFail = "Fail" Then
                            TEST3PASS = False
                            status("Red", "TEST3")
                            GlobalFail = GlobalFail + 1
                            TEST3Fail = TEST3Fail + 1
                            TEST3FailRetest = TEST3FailRetest + 1
                            UUTFail = 1
                        End If
                        Me.Failures3.Text = FormatPercent(((TEST3Fail / UUTNum)), 1)
                        Me.Total3.Text = UUTNum
                        Me.FailTotal3.Text = TEST3Fail

                        'AmplitudeBalance
                        'Directivity
                        If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And Not IsDBNull(drLocal.Item(11)) Then
                            RetrnVal = CDbl(drLocal.Item(11))
                        ElseIf (SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER") And Not IsDBNull(drLocal.Item(10)) Then
                            RetrnVal = CDbl(drLocal.Item(10))
                        Else
                            If Not TEST1PASS Then
                                Test1Fail = Test1Fail - 1
                                If Test1Fail = 0 Then status("Green", "TEST1")
                            End If
                            If Not TEST2PASS Then
                                status("Green", "TEST2")
                            End If
                            If Not TEST3PASS Then
                                status("Green", "TEST3")
                            End If
                            If Not TEST4PASS Then
                                status("Green", "TEST4")
                            End If
                            If Not Incomplete Then
                                If MsgBox("Data is incomplete at UUT" & UUTNum, vbOKOnly, "Resume Testing At " & UUTNum) = vbOK Then
                                    UUTNum = UUTNum - 1
                                    EraseTest.Text = "Remove UUT" & UUTNum
                                    UUTNum_Reset = UUTNum_Reset - 1
                                    Incomplete = True
                                    GoTo Recap
                                End If
                            End If
                        End If
                        PassFail = Tests.AmplitudeBalance(, True)
                        status("Blue", "TEST4")
                        RetrnVal = AB
                        PF4.Text = PassFail
                        RetrnStr = CStr(TruncateDecimal(RetrnVal, 2))
                        Data4.Text = Format(RetrnVal, "0.00")

                        If PassFail = "Pass" Then
                            TEST4PASS = True
                            status("Green", "TEST4")
                        ElseIf PassFail = "Fail" Then
                            TEST4PASS = False
                            status("Red", "TEST4")
                            GlobalFail = GlobalFail + 1
                            TEST4Fail = TEST4Fail + 1
                            TEST4FailRetest = TEST4FailRetest + 1
                            UUTFail = 1
                        End If
                        Me.Failures4.Text = FormatPercent(((TEST4Fail / UUTNum)), 1)
                        Me.Total4.Text = UUTNum
                        Me.FailTotal4.Text = TEST4Fail

                        'PhaseBalance
                        'Coupled Flatness
                        If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And Not IsDBNull(drLocal.Item(13)) Then
                            RetrnVal = CDbl(drLocal.Item(13))
                        ElseIf (SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER") And Not IsDBNull(drLocal.Item(12)) Then
                            RetrnVal = CDbl(drLocal.Item(12))
                        Else
                            If Not TEST1PASS Then
                                status("Green", "TEST1")
                            End If
                            If Not TEST2PASS Then
                                status("Green", "TEST2")
                            End If
                            If Not TEST3PASS Then
                                status("Green", "TEST3")
                            End If
                            If Not TEST4PASS Then
                                status("Green", "TEST4")
                            End If
                            If Not TEST5PASS Then
                                status("Green", "TEST5")
                            End If
                            If Not Incomplete Then
                                If MsgBox("Data is incomplete at UUT" & UUTNum, vbOKOnly, "Resume Testing At " & UUTNum) = vbOK Then
                                    UUTNum = UUTNum - 1
                                    EraseTest.Text = "Remove UUT" & UUTNum
                                    UUTNum_Reset = UUTNum_Reset - 1
                                    Incomplete = True
                                    GoTo Recap
                                End If
                            End If
                        End If
                        PassFail = Tests.PhaseBalance(SpecType, RetrnVal, True)

                        status("Blue", "TEST5")
                        PF5.Text = PassFail
                        RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                        Data5.Text = Format(RetrnVal, "0.0")
                        If PassFail = "Pass" Then
                            TEST5PASS = True
                            status("Green", "TEST5")
                        ElseIf PassFail = "Fail" Then
                            TEST5PASS = False
                            status("Red", "TEST5")
                            GlobalFail = GlobalFail + 1
                            TEST5Fail = TEST5Fail + 1
                            TEST5FailRetest = TEST5FailRetest + 1
                            UUTFail = 1
                        End If
                        Me.Failures5.Text = FormatPercent(((TEST5Fail / UUTNum)), 1)
                        Me.Total5.Text = UUTNum
                        Me.FailTotal5.Text = TEST5Fail

                        If Not TEST2PASS Or Not TEST1PASS Or Not TEST3PASS Or Not TEST4PASS Or Not TEST5PASS Then
                            UpdateFailureLog(Me.PF1.Text, Me.PF2.Text, Me.PF3.Text, Me.PF4.Text, Me.PF5.Text)
                            LOTFail = LOTFail + 1
                        End If
Recap1:
                        If LOTFail = 0 Then LotTestFrame.BackColor = Color.LawnGreen
                        If LOTFail = 0 Then LotFailureFrame.BackColor = Color.LawnGreen
                        If LOTFail > 0 Then LotTestFrame.BackColor = Color.Red
                        If LOTFail > 0 Then LotFailureFrame.BackColor = Color.Red
                        If LOTFail > 0 Then UUTStatusColor.BackColor = Color.Red
                        If UUTFail = 1 Then Me.cmdRetest.Enabled = True
                        If UUTFail = 0 Then Me.cmdRetest.Enabled = False
                        If UUTFail = 1 Then Me.EraseTest.Enabled = True
                        If UUTFail = 0 Then Me.EraseTest.Enabled = False
                        If UUTNum <> 0 Then Me.LotFailure.Text = FormatPercent(((LOTFail / UUTNum)), 1)
                        If TweakMode Then
                            UUTMessage.Text = "  UUT TESTS  --   Tweak Mode. No Data Logging"
                        ElseIf Not TraceChecked Then
                            UUTMessage.Text = "  UUT TESTS Marker Mode  --   Load Unit #" & UUTNum + 1
                        Else
                            UUTMessage.Text = "  UUT TESTS  --   Load Unit #" & UUTNum + 1
                        End If

                        Me.cmdStartTest.Text = "Next UUT"
                        cmdStartTest.Enabled = True
                        ' Me.cmdRetest.Enabled = True

                        MSChart.UpDateChart(SpecType)

                        TEST1PASS = True
                        TEST2PASS = True
                        TEST3PASS = True
                        TEST4PASS = True
                        TEST5PASS = True
                    End While
                End If
                UUTCount.Text = Str(UUTNum)
            End If


NoResume:
            'txtNet.Text = "Local"
            'txtNet.ForeColor = Color.DarkOrange
            txtTitle.Text = OldTitle
            Exit Function
Trap:
            MsgBox("Database Error. Cannot Resume")
            txtNet.Text = "Local"
            txtNet.ForeColor = Color.DarkOrange
            txtTitle.Text = OldTitle

        Catch ex As Exception
            ResumeTesting = False
        End Try
    End Function


    Private Sub txtOffset_DblClick(Index As Integer)
        Dim SQLstr As String

        Dim Test
        On Error GoTo Trap
        If Index = 1 Then If Not IsNumeric(Me.txtOffset1.Text) Then Exit Sub
        If Index = 2 Then If Not IsNumeric(Me.txtOffset2.Text) Then Exit Sub
        If Index = 3 Then If Not IsNumeric(Me.txtOffset3.Text) Then Exit Sub
        If Index = 4 Then If Not IsNumeric(Me.txtOffset4.Text) Then Exit Sub
        If Index = 5 Then If Not IsNumeric(Me.txtOffset5.Text) Then Exit Sub
        SQLstr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
        If SQL.CheckforRow(SQLstr, "LocalSpecs") = 0 Then
            SQLstr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
            SQL.ExecuteSQLCommand(SQLstr, "LocalSpecs")
        End If
        If SQLAccess Then
            Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
            Dim cmd As SqlCommand = New SqlCommand()
            ats.Open()
            cmd.Connection = ats
            cmd.CommandText = "UPDATE Specifications Set Offset1 = '" & Me.txtOffset1.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "UPDATE Specifications Set Offset2 = '" & Me.txtOffset2.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "UPDATE Specifications Set Offset3 = '" & Me.txtOffset3.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "UPDATE Specifications Set Offset4 = '" & Me.txtOffset4.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "UPDATE Specifications Set Offset5 = '" & Me.txtOffset5.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            cmd.ExecuteNonQuery()
            ats.Close()
        Else
            Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkData")
            Dim atsLocal As New OleDb.OleDbConnection
            atsLocal.ConnectionString = strConnectionString
            Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
            atsLocal.Open()
            cmd.Connection = atsLocal
            cmd.CommandText = "UPDATE Specifications Set Offset1 = '" & Me.txtOffset1.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "UPDATE Specifications Set Offset2 = '" & Me.txtOffset2.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "UPDATE Specifications Set Offset3 = '" & Me.txtOffset3.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "UPDATE Specifications Set Offset4 = '" & Me.txtOffset4.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "UPDATE Specifications Set Offset5 = '" & Me.txtOffset5.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            cmd.ExecuteNonQuery()
            atsLocal.Close()
        End If

        Exit Sub
Trap:

    End Sub

    Private Sub txtOffset_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        Dim SQLstr As String
        On Error GoTo Trap

        If Index = 13 Then
            SQLstr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
            If SQL.CheckforRow(SQLstr, "LocalSpecs") = 0 Then
                SQLstr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
                SQL.ExecuteSQLCommand(SQLstr, "LocalSpecs")
            End If
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE Specifications Set Offset1 = '" & Me.txtOffset1.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Specifications Set Offset2 = '" & Me.txtOffset2.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Specifications Set Offset3 = '" & Me.txtOffset3.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Specifications Set Offset4 = '" & Me.txtOffset4.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Specifications Set Offset5 = '" & Me.txtOffset5.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.Connection = atsLocal
                cmd.CommandText = "UPDATE Specifications Set Offset1 = '" & Me.txtOffset1.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Specifications Set Offset2 = '" & Me.txtOffset2.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Specifications Set Offset3 = '" & Me.txtOffset3.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Specifications Set Offset4 = '" & Me.txtOffset4.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Specifications Set Offset5 = '" & Me.txtOffset5.Text & "' where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If
        End If

        Exit Sub
Trap:
    End Sub


    Private Sub UploadTrace()
        Dim SQLstr As String
        Dim LCL As ADODB.Recordset
        Dim Net As ADODB.Recordset
        Dim Trace As ADODB.Recordset
        Dim POINTSNET As ADODB.Recordset
        Dim POINTSLOCAL As ADODB.Recordset
        Dim i, j As Integer
        Dim MessageTxt As String
        Dim TempAccess As Boolean
        Dim Count As Integer
        Dim XData(1, 1)
        Dim YData(1, 1)
        Dim TraceID(1, 1)
        Dim Idx(1, 1)

        On Error GoTo Trap

        txtNet.Text = "Network"
        txtNet.ForeColor = Color.LawnGreen
        MessageTxt = Me.txtTitle.Text
        Me.txtTitle.Text = "    PLEASE WAIT....... UPLOADING TRACE DATA TO SQL Server...........  "
        TempAccess = SQLAccess

        ' Local Trace Data
        SQLAccess = False
        SQLstr = "select * from Trace"
        Count = SQL.CheckforRow(SQLstr, "LocalTraceData")
        Dim JobNumber(Count) As String
        Dim Title(Count) As String
        Dim SerialNumber(Count) As String
        Dim WorkStation(Count) As String
        Dim TestID(Count) As String
        Dim SpecID(Count) As String
        Dim Workstation1(Count) As String
        Dim Points(Count) As Integer
        Dim ActiveDateArray(Count) As String
        Dim RFPower(Count) As String
        Dim Temperature(Count) As String
        Dim CalibrationDate(Count) As String
        Dim InstrumentCalDue(Count) As String
        Dim ProgTitle(Count) As String
        Dim ProgVer(Count) As String
        Dim XTitle(Count) As String
        Dim Ytitle(Count) As String
        Dim Notes(Count) As String

        ExpectedProgress.Minimum = 0
        SQLstr = "select * from Trace"
        Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkSpecs")
        Dim atsLocal As New OleDb.OleDbConnection
        atsLocal.ConnectionString = strConnectionString
        Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
        atsLocal.Open()
        System.Threading.Thread.Sleep(10)
        Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
        i = 0
        While Not drLocal.Read = Nothing
            i = i + 1
            TestID(i) = CStr(drLocal.Item(1))
            SpecID(i) = CStr(drLocal.Item(2))
            JobNumber(i) = CStr(drLocal.Item(3))
            Title(i) = CStr(drLocal.Item(4))
            SerialNumber(i) = CStr(drLocal.Item(5))
            Workstation1(i) = CStr(drLocal.Item(6))
            Points(i) = CInt(drLocal.Item(7))
            ActiveDateArray(i) = CStr(drLocal.Item(8))
            RFPower(i) = CStr(drLocal.Item(9))
            Temperature(i) = CStr(drLocal.Item(10))
            CalibrationDate(i) = CStr(drLocal.Item(11))
            InstrumentCalDue(i) = CStr(drLocal.Item(12))
            ProgTitle(i) = CStr(drLocal.Item(13))
            ProgVer(i) = CStr(drLocal.Item(14))
            XTitle(i) = CStr(drLocal.Item(15))
            Ytitle(i) = CStr(drLocal.Item(16))
            Notes(i) = CStr(drLocal.Item(17))


            If Count <> 0 Then ExpectedProgress.Maximum = ((Count) * (Points(i) + 1))
            Me.txtTitle.Text = "..........  UPLOADING TRACE " & i & " OF " & Count & " TO SQL SERVER............  "
            System.Threading.Thread.Sleep(10)

            '*****************Trace Points********************************
            ReDim XData(Count, Points(i))
            ReDim YData(Count, Points(i))
            ReDim TraceID(Count, Points(i))
            ReDim Idx(Count, Points(i))
            SQLstr = "select * from TracePoints where TraceID = '" & drLocal.Item(0) & "'"
            Dim atsLocal1 As New OleDb.OleDbConnection
            atsLocal1.ConnectionString = strConnectionString
            Dim cmd1 As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
            atsLocal1.Open()
            System.Threading.Thread.Sleep(10)
            Dim drLocal1 As OleDb.OleDbDataReader = cmd.ExecuteReader
            j = 0
            While Not drLocal1.Read = Nothing
                j = j + 1
                XData(1, j) = CStr(drLocal1.Item(1))
                YData(1, j) = CStr(drLocal1.Item(1))
                TraceID(1, j) = CStr(drLocal1.Item(1))
                Idx(1, j) = CStr(drLocal1.Item(1))
            End While
            drLocal1.Close()
            If ExpectedProgress.Value + 1 < ExpectedProgress.Maximum Then ExpectedProgress.Value = ExpectedProgress.Value + 0.5
        End While
        drLocal.Close()


        ' NetworkTrace Data
        SQLAccess = True
        Dim ats2 As SqlConnection = New SqlConnection(SQLConnStr)
        Dim cmd2 As SqlCommand = New SqlCommand()
        ats2.Open()
        cmd2.Connection = ats2
        For i = 0 To Count
            SQLstr = "select * from Trace where TestID = '" & TestID(i) & "'"
            If SQL.CheckforRow(SQLstr, "LocalTestData") = 0 Then
                SQLstr = "Insert Into Trace (TestID) values ('" & TestID(i) & "')"
                SQL.ExecuteSQLCommand(SQLstr, "LocalSpecs")
            End If
            cmd2.CommandText = "UPDATE Trace Set SpecID = '" & SpecID(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set JobNumber = '" & JobNumber(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set Title(= '" & Title(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set SerialNumber = '" & SerialNumber(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set Workstation1 = '" & Workstation1(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set ActiveDate = '" & ActiveDateArray(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set RFPower = '" & RFPower(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set Temperature = '" & Temperature(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set CalibrationDate = '" & CalibrationDate(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set InstrumentCalDue = '" & InstrumentCalDue(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set ProgTitle = '" & ProgTitle(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set ProgVer = '" & ProgVer(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set XTitle = '" & XTitle(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set Ytitle = '" & Ytitle(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            cmd2.CommandText = "UPDATE Trace Set Notes = '" & Notes(i) & "' where TestID = " & TestID(i)
            cmd2.ExecuteNonQuery()
            Dim ats3 As SqlConnection = New SqlConnection(SQLConnStr)
            Dim cmd3 As SqlCommand = New SqlCommand()
            ats3.Open()
            cmd3.Connection = ats3
            For Y = 0 To Points(i)
                SQLstr = "select * from TracePoints where TestID = '" & TestID(i) & "'"
                If SQL.CheckforRow(SQLstr, "LocalTraceData") = 0 Then
                    SQLstr = "Insert Into Points (TestID) values ('" & TestID(i) & "')"
                    SQL.ExecuteSQLCommand(SQLstr, "LocalTraceData")
                End If
                cmd3.CommandText = "UPDATE TracePoints Set XData = '" & XData(1, j) & "' where TestID = " & TestID(i)
                cmd3.ExecuteNonQuery()
                cmd3.CommandText = "UPDATE TracePoints Set YData = '" & YData(1, j) & "' where TestID = " & TestID(i)
                cmd3.ExecuteNonQuery()
                cmd3.CommandText = "UPDATE TracePoints Set TraceID = '" & TraceID(1, j) & "' where TestID = " & TestID(i)
                cmd3.ExecuteNonQuery()
                cmd3.CommandText = "UPDATE TracePoints Set Idx = '" & Idx(1, j) & "' where TestID = " & TestID(i)
                cmd3.ExecuteNonQuery()
            Next
            ats3.Close()
        Next
        ats2.Close()


        Me.txtTitle.Text = "    ............DELETING LOCAL Trace Data............  "
        System.Threading.Thread.Sleep(100)
        SQLstr = "Delete from Trace"
        SQL.ExecuteSQLCommand(SQLstr, "LocalTraceData")


        SQLstr = "Delete from TracePoints"
        SQL.ExecuteSQLCommand(SQLstr, "LocalTraceData")


        '    Me.MasterUpLoad.Enabled = False
        Me.txtTitle.Text = MessageTxt
        txtNet.Text = "Network"
        txtNet.ForeColor = Color.LawnGreen
        Exit Sub
Trap:
        txtNet.Text = "Local"
        txtNet.ForeColor = Color.DarkOrange
    End Sub

    Private Sub UploadTestData(DataLocation As String)
        Dim SQLstr As String
        Dim i, Count As Integer
        Dim MessageTxt As String
        Dim OldSQLACCESS As Boolean

        Try
            txtNet.Text = "Network"
            txtNet.ForeColor = Color.LawnGreen
            Me.Refresh()
            MessageTxt = Me.txtTitle.Text
            Me.txtTitle.Text = "    PLEASE WAIT....... UPLOADING LOCAL TEST DATA TO SQL ...........  "
            ' Test Data
            Me.Refresh()
            SQLstr = "Select * from TestData Where JobNumber = '" & Job & "'"
            OldSQLACCESS = SQLAccess
            SQLAccess = False
            Count = SQL.CheckforRow(SQLstr, DataLocation)
            If Count = 0 Then
                OldSQLACCESS = SQLAccess
                Me.txtTitle.Text = MessageTxt
                If SQLAccess Then
                    txtNet.Checked = True
                    txtNet.Text = "SQL Database"
                ElseIf NetworkAccess Then
                    txtNet.Text = "Network Access Database"
                Else
                    txtNet.Text = "Local Access Database"
                End If
                Exit Sub
            End If
            Dim TestID(Count) As String
            Dim SpecID(Count) As String
            Dim SerialNumber(Count) As String
            Dim Workstation(Count) As String
            Dim InsertionLoss(Count) As Double
            Dim ReturnLoss(Count) As Double
            Dim Coupling(Count) As Double
            Dim Isolation(Count) As Double
            Dim Directivity(Count) As Double
            Dim CoupledFlatness(Count) As Integer

            Dim AmplitudeBalance(Count) As Double
            Dim PhaseBalance(Count) As Double
            Dim FailureLog(Count) As String
            ExpectedProgress.Minimum = 0


            SQLstr = "select * from TestData Where JobNumber = '" & Job & "'"
            Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder(DataLocation)
            Dim atsLocal As New OleDb.OleDbConnection
            atsLocal.ConnectionString = strConnectionString
            Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
            atsLocal.Open()
            System.Threading.Thread.Sleep(0.001)
            Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
            i = 0
            While Not drLocal.Read = Nothing
                If Not IsDBNull(drLocal.Item(0)) Then
                    TestID(i) = drLocal.Item(0)
                Else
                    TestID(i) = ""
                End If
                If Not IsDBNull(drLocal.Item(1)) Then
                    SpecID(i) = CStr(drLocal.Item(1))
                Else
                    SpecID(i) = ""
                End If
                If Not IsDBNull(drLocal.Item(4)) Then
                    SerialNumber(i) = CStr(drLocal.Item(4))
                Else
                    SpecID(i) = ""
                End If

                If Not IsDBNull(drLocal.Item(5)) Then
                    Workstation(i) = CStr(drLocal.Item(5))
                Else
                    Workstation(i) = ""
                End If
                If Not IsDBNull(drLocal.Item(6)) Then
                    InsertionLoss(i) = CDbl(drLocal.Item(6))
                Else
                    InsertionLoss(i) = ""
                End If
                If Not IsDBNull(drLocal.Item(7)) Then
                    ReturnLoss(i) = CDbl(drLocal.Item(7))
                Else
                    ReturnLoss(i) = "0"
                End If
                If Not IsDBNull(drLocal.Item(8)) Then
                    Coupling(i) = CStr(drLocal.Item(8))
                Else
                    Coupling(i) = "0"
                End If
                If Not IsDBNull(drLocal.Item(9)) Then
                    Isolation(i) = CDbl(drLocal.Item(9))
                Else
                    Isolation(i) = "0"
                End If
                If Not IsDBNull(drLocal.Item(10)) Then
                    Directivity(i) = CDbl(drLocal.Item(10))
                Else
                    Directivity(i) = "0"
                End If
                If Not IsDBNull(drLocal.Item(11)) Then
                    AmplitudeBalance(i) = CDbl(drLocal.Item(11))
                Else
                    AmplitudeBalance(i) = "0"
                End If
                If Not IsDBNull(drLocal.Item(12)) Then
                    CoupledFlatness(i) = CInt(drLocal.Item(12))
                Else
                    CoupledFlatness(i) = "0"
                End If

                If Not IsDBNull(drLocal.Item(13)) Then
                    PhaseBalance(i) = CDbl(drLocal.Item(13))
                Else
                    PhaseBalance(i) = "0"
                End If
                If Not IsDBNull(drLocal.Item(14)) Then
                    FailureLog(i) = CStr(drLocal.Item(14))
                Else
                    FailureLog(i) = ""
                End If
                i = i + 1
            End While
            drLocal.Close()
            ' NetworkTrace Data
            SQLAccess = True
            Dim ats2 As SqlConnection = New SqlConnection(SQLConnStr)
            Dim cmd2 As SqlCommand = New SqlCommand()
            ats2.Open()
            cmd2.Connection = ats2
            For i = 0 To Count - 1
                SQLstr = "select * from TestData where SpecID = '" & TestID(i) & "'"
                If SQL.CheckforRow(SQLstr, "NetworkData") = 0 Then
                    SQLstr = "Insert Into TestData (SpecID) values ('" & TestID(i) & "')"
                    SQL.ExecuteSQLCommand(SQLstr, "NetworkData")
                End If
                cmd2.CommandText = "UPDATE TestData Set JobNumber = '" & Me.cmbJob.Text & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set PartNumber = '" & Me.cmbPart.Text & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set SerialNumber = '" & SerialNumber(i) & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set Workstation = '" & Workstation(i) & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set InsertionLoss = '" & InsertionLoss(i) & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set ReturnLoss = '" & ReturnLoss(i) & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set Coupling = '" & Coupling(i) & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set Directivity = '" & Directivity(i) & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set CoupledFlatness = '" & CoupledFlatness(i) & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set Isolation = '" & Isolation(i) & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set AmplitudeBalance = '" & AmplitudeBalance(i) & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set PhaseBalance = '" & PhaseBalance(i) & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
                cmd2.CommandText = "UPDATE TestData Set FailureLog = '" & FailureLog(i) & "' where SpecID = " & TestID(i)
                cmd2.ExecuteNonQuery()
            Next
            ats2.Close()
            Me.txtTitle.Text = MessageTxt
            If SQLAccess Then
                txtNet.Checked = True
                txtNet.Text = "SQL Database"
            ElseIf NetworkAccess Then
                txtNet.Text = "Network Access Database"

            Else
                txtNet.Text = "Local Access Database"
            End If

            txtNet.ForeColor = Color.LawnGreen
            Exit Sub
        Catch
            OldSQLACCESS = SQLAccess
            Me.txtTitle.Text = "ERROR"
            If SQLAccess Then
                txtNet.Checked = True
                txtNet.Text = "SQL Database"
            ElseIf NetworkAccess Then
                txtNet.Text = "Network Access Database"
            Else
                txtNet.Text = "Local Access Database"
            End If

        End Try
    End Sub

    Private Sub UploadSpecData()
        '        Dim SQLstr As String
        '        Dim LCL As ADODB.Recordset
        '        Dim Net As ADODB.Recordset
        '        Dim Trace As ADODB.Recordset
        '        Dim POINTSNET As ADODB.Recordset
        '        Dim POINTSLOCAL As ADODB.Recordset
        '        Dim i, j As Integer
        '        Dim MessageTxt As String
        'On Errror GoTo Trap

        '        txtNet.Text = "Network"
        '        txtNet.ForeColor = Color.LawnGreen
        '        MessageTxt = Me.txtTitle.Text
        '        Me.txtTitle.Text = "    PLEASE WAIT....... UPLOADING Spec DATA TO THE NETWORK............  "
        '        ' Test Data
        '        SQLstr = "select * from Specficiations"
        '        LCL = New ADODB.Recordset
        '        LCL.Open(SQLstr, NetSpecsConn, adOpenStatic, adLockBatchOptimistic)
        '        ' ExpectedProgress.Min = 0
        '        If LCL.RecordCount <= 1 Then
        '            ' ExpectedProgress.Max = 1
        '        Else
        '            ' ExpectedProgress.Max = Val(LCL.RecordCount) - 1
        '        End If

        '        ' If LCL.RecordCount = 0 Then ExpectedProgress.Value = 1
        '        For i = 0 To LCL.RecordCount - 1
        '            Me.txtTitle.Text = "..........  UPLOADING RECORD " & i & " OF " & LCL.RecordCount & " TO SQL SERVER...........  "
        '            SQLstr = "select * from Specifications where JobNumber = '" & LCL!JobNumber & "' and PartNumber = '" & LCL!PartNumber & "'"
        '            Net = New ADODB.Recordset
        '            Net.Open(SQLstr, NetConn, adOpenDynamic, adLockOptimistic)
        '            If Net.EOF Then
        '                SQLstr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & LCL!JobNumber & "','" & LCL!PartNumber & "')"
        '                NetConn.Execute SQLstr
        '                Delay 5

        '                SQLstr = "select * from Specifications where JobNumber = '" & LCL!JobNumber & "' and PartNumber = '" & LCL!PartNumber & "'"
        '                Net = New ADODB.Recordset
        '                Net.Open(SQLstr, NetConn, adOpenDynamic, adLockOptimistic)
        '            End If
        '            If Not IsNull(LCL!SpecType) Then Net!SpecType = LCL!SpecType
        '            If Not IsNull(LCL!JobNumber) Then Net!JobNumber = LCL!JobNumber
        '            If Not IsNull(LCL!PartNumber) Then Net!PartNumber = LCL!PartNumber
        '            If Not IsNull(LCL!Title) Then Net!Title = LCL!Title
        '            If Not IsNull(LCL!Quantity) Then Net!Quantity = LCL!Quantity
        '            If Not IsNull(LCL!StartFreqMHz) Then Net!StartFreqMHz = LCL!StartFreqMHz
        '            If Not IsNull(LCL!StopFreqMHz) Then Net!StopFreqMHz = LCL!StopFreqMHz
        '            If Not IsNull(LCL!CutOffFreqMHz) Then Net!CutOffFreqMHz = LCL!CutOffFreqMHz
        '            If Not IsNull(LCL!OutputPortNumber) Then Net!OutputPortNumber = LCL!OutputPortNumber
        '            If Not IsNull(LCL!VSWR) Then Net!VSWR = LCL!VSWR
        '            If Not IsNull(LCL!InsertionLoss) Then Net!InsertionLoss = LCL!InsertionLoss
        '            If Not IsNull(LCL!Isolation) Then Net!Isolation = LCL!Isolation
        '            If Not IsNull(LCL!Isolation2) Then Net!Isolation2 = LCL!Isolation2
        '            If Not IsNull(LCL!AmplitudeBalance) Then Net!AmplitudeBalance = LCL!AmplitudeBalance
        '            If Not IsNull(LCL!Coupling) Then Net!Coupling = LCL!Coupling
        '            If Not IsNull(LCL!Directivity) Then Net!Directivity = LCL!Directivity
        '            If Not IsNull(LCL!PhaseBalance) Then Net!PhaseBalance = LCL!PhaseBalance
        '            If Not IsNull(LCL!CoupledFlatness) Then Net!CoupledFlatness = LCL!CoupledFlatness
        '            If Not IsNull(LCL!Power) Then Net!Power = LCL!Power
        '            If Not IsNull(LCL!Temperature) Then Net!Temperature = LCL!Temperature
        '            If Not IsNull(LCL!Offset1) Then Net!Offset1 = LCL!Offset1
        '            If Not IsNull(LCL!Offset2) Then Net!Offset2 = LCL!Offset2
        '            If Not IsNull(LCL!Offset3) Then Net!Offset3 = LCL!Offset3
        '            If Not IsNull(LCL!Offset4) Then Net!Offset4 = LCL!Offset4
        '            If Not IsNull(LCL!Offset5) Then Net!Offset5 = LCL!Offset5
        '            If Not IsNull(LCL!Test1) Then Net!Test1 = LCL!Test1
        '            If Not IsNull(LCL!TEST2) Then Net!TEST2 = LCL!TEST2
        '            If Not IsNull(LCL!TEST3) Then Net!TEST3 = LCL!TEST3
        '            If Not IsNull(LCL!TEST4) Then Net!TEST4 = LCL!TEST4
        '            If Not IsNull(LCL!TEST5) Then Net!TEST5 = LCL!TEST5

        '            Net.Update()
        '            Net.Close()
        '            LCL.MoveNext()
        '            'ExpectedProgress.Value = i
        '        Next i
        '        LCL.Close()


        '        Me.txtTitle.Text = MessageTxt
        'If SQLAccess Then
        '    txtNet.Text = "SQL Database"
        'ElseIf NetworkAccess Then
        '    txtNet.Text = "Network Access Database"
        'Else
        '    txtNet.Text = "Local Access Database"
        'End If
        '        txtNet.ForeColor = Color.LawnGreen
        '        Exit Function
        'Trap:
        '        txtNet.Text = "Local"
        '        txtNet.ForeColor = DarkOrange
    End Sub

    Private Sub UploadPortConfig()
        '        Dim SQLstr As String
        '        Dim LCL As ADODB.Recordset
        '        Dim Net As ADODB.Recordset
        '        Dim Trace As ADODB.Recordset
        '        Dim POINTSNET As ADODB.Recordset
        '        Dim POINTSLOCAL As ADODB.Recordset
        '        Dim i, j As Integer
        '        Dim MessageTxt As String
        'On Errror GoTo Trap

        '        txtNet.Text = "Network"
        '        txtNet.ForeColor = Color.LawnGreen
        '        MessageTxt = Me.txtTitle.Text
        '        Me.txtTitle.Text = "    PLEASE WAIT....... UPLOADING Spec DATA TO THE NETWORK............  "
        '        ' Test Data
        '        SQLstr = "select * from PortConfig"
        '        LCL = New ADODB.Recordset
        '        LCL.Open(SQLstr, NetSpecsConn, adOpenStatic, adLockBatchOptimistic)
        '        ' ExpectedProgress.Min = 0
        '        If LCL.RecordCount <= 1 Then
        '            ' ExpectedProgress.Max = 1
        '        Else
        '            ' ExpectedProgress.Max = Val(LCL.RecordCount) - 1
        '        End If

        '        'If LCL.RecordCount = 0 Then ExpectedProgress.Value = 1
        '        For i = 0 To LCL.RecordCount - 1
        '            Me.txtTitle.Text = "..........  UPLOADING RECORD " & i & " OF " & LCL.RecordCount & " TO SQL SERVER...........  "
        '            SQLstr = "select * from  PortConfig where JobNumber = '" & LCL!JobNumber & "' and PartNumber = '" & LCL!PartNumber & "'"
        '            Net = New ADODB.Recordset
        '            Net.Open(SQLstr, NetConn, adOpenDynamic, adLockOptimistic)
        '            If Net.EOF Then
        '                SQLstr = "Insert Into  PortConfig (JobNumber, PartNumber) values ('" & LCL!JobNumber & "','" & LCL!PartNumber & "')"
        '                NetConn.Execute SQLstr
        '                Delay 100

        '                SQLstr = "select * from  PortConfig where JobNumber = '" & LCL!JobNumber & "' and PartNumber = '" & LCL!PartNumber & "'"
        '                Net = New ADODB.Recordset
        '                Net.Open(SQLstr, NetConn, adOpenDynamic, adLockOptimistic)
        '            End If
        '            If Not IsNull(LCL!JobNumber) Then Net!JobNumber = LCL!JobNumber
        '            If Not IsNull(LCL!PartNumber) Then Net!PartNumber = LCL!PartNumber

        '            If Not IsNull(LCL!J1J1) Then Net!J1J1 = LCL!J1J1
        '            If Not IsNull(LCL!J1J2) Then Net!J1J2 = LCL!J1J2
        '            If Not IsNull(LCL!J1J3) Then Net!J1J3 = LCL!J1J3
        '            If Not IsNull(LCL!J1J4) Then Net!J1J4 = LCL!J1J4
        '            If Not IsNull(LCL!J1J5) Then Net!J1J5 = LCL!J1J5
        '            If Not IsNull(LCL!J2J1) Then Net!J2J1 = LCL!J2J1
        '            If Not IsNull(LCL!J2J2) Then Net!J2J2 = LCL!J2J2
        '            If Not IsNull(LCL!J2J3) Then Net!J2J3 = LCL!J3J3
        '            If Not IsNull(LCL!J2J4) Then Net!J2J4 = LCL!J2J4
        '            If Not IsNull(LCL!J3J1) Then Net!J3J1 = LCL!J3J1
        '            If Not IsNull(LCL!J3J2) Then Net!J3J2 = LCL!J3J2
        '            If Not IsNull(LCL!J3J3) Then Net!J3J3 = LCL!J3J3
        '            If Not IsNull(LCL!J3J4) Then Net!J3J4 = LCL!J3J4
        '            If Not IsNull(LCL!J4J1) Then Net!J4J1 = LCL!J4J1
        '            If Not IsNull(LCL!J4J2) Then Net!J4J2 = LCL!J4J2
        '            If Not IsNull(LCL!J4J3) Then Net!J4J3 = LCL!J4J3
        '            If Not IsNull(LCL!J4J4) Then Net!J4J4 = LCL!J4J4


        '            Net.Update()
        '            Net.Close()
        '            LCL.MoveNext()
        '            ' ExpectedProgress.Value = i
        '        Next i
        '        LCL.Close()


        '        Me.txtTitle.Text = MessageTxt
        'If SQLAccess Then
        '    txtNet.Text = "SQL Database"
        'ElseIf NetworkAccess Then
        '    txtNet.Text = "Network Access Database"
        'Else
        '    txtNet.Text = "Local Access Database"
        'End If
        '        txtNet.ForeColor = Color.LawnGreen
        '                '        Exit Function
        'Trap:
        '        txtNet.Text = "Local"
        '        txtNet.ForeColor = DarkOrange
        '        
    End Sub



    'Private Function MasterUploadData()
    'Dim SQLstr As String
    'Dim LCL As ADODB.Recordset
    'Dim Net As ADODB.Recordset
    'Dim Trace As ADODB.Recordset
    'Dim POINTSNET As ADODB.Recordset
    'Dim POINTSLOCAL As ADODB.Recordset
    'Dim i, j As Integer
    'Dim MessageTxt As String
    'On Errror GoTo Trap
    '
    '
    '
    '    txtNet.Text = "Network"
    '    txtNet.ForeColor = &HFFFF&
    '    Network.BackColor = &HFFFF&
    '    MessageTxt = Me.txtTitle.Text
    '    Me.txtTitle.Text = "    PLEASE WAIT....... UPLOADING TEST DATA TO THE NETWORK............  "
    '    ' Test Data
    '    SQLstr = "select * from TestData"
    '    Set LCL = New ADODB.Recordset
    '    LCL.Open SQLstr, NetworkDataConn, adOpenStatic, adLockBatchOptimistic
    '   ' ExpectedProgress.Min = 0
    '    If LCL.RecordCount <= 1 Then
    '      '  ExpectedProgress.Max = 1
    '    Else
    '        'ExpectedProgress.Max = Val(LCL.RecordCount) - 1
    '    End If
    '
    '    'If LCL.RecordCount = 0 Then ExpectedProgress.Value = 1
    '    For i = 0 To LCL.RecordCount - 1
    '        Me.txtTitle.Text = "..........  UPLOADING RECORD " & i & " OF " & LCL.RecordCount & " TO THE NETWORK............  "
    '        SQLstr = "select * from TestData"
    '        Set Net = New ADODB.Recordset
    '        Net.Open SQLstr, NetDataConn, adOpenDynamic, adLockOptimistic
    '        If Net.EOF Then
    '            SQLstr = "Insert Into TestData(JobNumber, PartNumber, SerialNumber) values ('" & LCL!JobNumber & "','" & LCL!PartNumber & "','" & LCL!SerialNumber & "')"
    '            NetDataConn.Execute SQLstr
    '            Delay 100
    '
    '            SQLstr = "select * from TestData where JobNumber = '" & LCL!JobNumber & "' and PartNumber = '" & LCL!PartNumber & "' and SerialNumber = '" & LCL!SerialNumber & "'"
    '            Set Net = New ADODB.Recordset
    '            Net.Open SQLstr, NetDataConn, adOpenDynamic, adLockOptimistic
    '        End If
    '
    '        If Not IsNull(LCL!SpecID) Then Net!SpecID = LCL!SpecID
    '        If Not IsNull(LCL!Workstation) Then Net!Workstation = LCL!Workstation
    '        If Not IsNull(LCL!InsertionLoss) Then Net!InsertionLoss = LCL!InsertionLoss
    '        If Not IsNull(LCL!ReturnLoss) Then Net!ReturnLoss = LCL!ReturnLoss
    '        If InStr(SpecType, "DIRECTIONAL COUPLER") Then
    '            If Not IsNull(LCL!Coupling) Then Net!Coupling = LCL!Coupling
    '            If Not IsNull(LCL!Directivity) Then Net!Directivity = LCL!Directivity
    '            If Not IsNull(LCL!CoupledFlatness) Then Net!CoupledFlatness = LCL!CoupledFlatness
    '        Else
    '            If Not IsNull(LCL!Isolation) Then Net!Isolation = LCL!Isolation
    '            If Not IsNull(LCL!AmplitudeBalance) Then Net!AmplitudeBalance = LCL!AmplitudeBalance
    '            If Not IsNull(LCL!PhaseBalance) Then Net!PhaseBalance = LCL!PhaseBalance
    '        End If
    '
    '
    '        If Not IsNull(LCL!FailureLog) Then Net!FailureLog = LCL!FailureLog
    '        Net.Update
    '        Net.Close
    '        LCL.MoveNext
    '       ' ExpectedProgress.Value = i
    '    Next i
    '    LCL.Close
    '    Me.txtTitle.Text = "    ............DELETING LOCAL RECORDS............  "
    '
    '    SQLstr = "Delete from TestData"
    '    Set LCL = New ADODB.Recordset
    '    LCL.Open SQLstr, LocalDataConn, adOpenDynamic, adLockOptimistic
    '
    '
    '    Me.txtTitle.Text = "    PLEASE WAIT....... UPLOADING TRACE Points TO THE NETWORK ............  "
    '    ' Trace Data
    '    SQLstr = "select * from Trace"
    '    Set LCL = New ADODB.Recordset
    '    LCL.Open SQLstr, NetTraceConn, adOpenStatic, adLockBatchOptimistic
    '
    '
    '  '  ExpectedProgress.Min = 0
    '   ' If LCL.RecordCount <> 0 Then ExpectedProgress.Max = ((LCL.RecordCount) * (Points + 1))
    '    For i = 0 To LCL.RecordCount - 1
    '        Me.txtTitle.Text = "..........  UPLOADING TRACE " & i & " OF " & LCL.RecordCount & " TO THE NETWORK............  "
    '        SQLstr = "select * from TracePoints where TraceID = " & LCL!ID & ""
    '        Set POINTSLOCAL = New ADODB.Recordset
    '        POINTSLOCAL.Open SQLstr, LocalTraceConn, adOpenStatic, adLockBatchOptimistic
    '
    '        For j = 0 To POINTSLOCAL.RecordCount - 1
    '            Me.txtTitle.Text = "    PLEASE WAIT....... UPLOADING " & LCL.RecordCount & " TRACES:  Trace " & i & "  TO THE NETWORK ............  "
    '            SQLstr = "select * from TracePoints where TraceID = " & LCL!ID & " and Idx = " & j & ""
    '            Set POINTSNET = New ADODB.Recordset
    '            POINTSNET.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    '            If POINTSNET.EOF Then
    '                SQLstr = "Insert Into TracePoints(TraceID, Idx) values ('" & LCL!ID & "','" & j & "')"
    '                NetTraceConn.Execute SQLstr
    '                Delay 100
    '
    '                SQLstr = "select * from TracePoints where TraceID = " & LCL!ID & " and Idx = " & j & ""
    '                Set POINTSNET = New ADODB.Recordset
    '                POINTSNET.Open SQLstr, NetTraceConn, adOpenDynamic, adLockOptimistic
    '            End If
    '            If Not IsNull(POINTSLOCAL!XData) Then POINTSNET!XData = POINTSLOCAL!XData
    '            If Not IsNull(POINTSLOCAL!YData) Then POINTSNET!YData = POINTSLOCAL!YData
    '            POINTSNET.Update
    '            POINTSNET.Close
    '            POINTSLOCAL.MoveNext
    '          '  If ExpectedProgress.Value + 1 < ExpectedProgress.Max Then ExpectedProgress.Value = ExpectedProgress.Value + 1
    '         Next j
    '
    '         SQLstr = "select * from Trace where JobNumber = '" & LCL!JobNumber & "' and Title = '" & LCL!Title & "' and SerialNumber = '" & LCL!SerialNumber & "'"
    '         Set Net = New ADODB.Recordset
    '         Net.Open SQLstr, NetConn, adOpenStatic, adLockOptimistic
    '         If Net.EOF Then
    '            Net.Close
    '            SQLstr = "Insert Into Trace(JobNumber, Title, SerialNumber) values ('" & LCL!JobNumber & "','" & LCL!Title & "','" & LCL!SerialNumber & "')"
    '            NetTraceConn.Execute SQLstr
    '            Delay 100
    '
    '            SQLstr = "select * from Trace where JobNumber = '" & LCL!JobNumber & "' and Title = '" & LCL!Title & "' and SerialNumber = '" & LCL!SerialNumber & "'"
    '            Set POINTSNET = New ADODB.Recordset
    '            Net.Open SQLstr, NetTraceConn, adOpenDynamic, adLockOptimistic
    '         End If
    '        If Not IsNull(LCL!TestID) Then Net!TestID = LCL!TestID
    '        If Not IsNull(LCL!SpecID) Then Net!SpecID = LCL!SpecID
    '        If Not IsNull(LCL!Workstation) Then Net!Workstation = LCL!Workstation
    '        If Not IsNull(LCL!Points) Then Net!Points = LCL!Points
    '        If Not IsNull(LCL!ActiveDate) Then Net!ActiveDateArray( = LCL!ActiveDate
    '        If Not IsNull(LCL!RFPower) Then Net!RFPower = LCL!RFPower
    '        If Not IsNull(LCL!Temperature) Then Net!Temperature = LCL!Temperature
    '        If Not IsNull(LCL!CalibrationDate) Then Net!CalibrationDate = LCL!CalibrationDate
    '        If Not IsNull(LCL!InstrumentCalDue) Then Net!InstrumentCalDue = LCL!InstrumentCalDue
    '        If Not IsNull(LCL!ProgTitle) Then Net!ProgTitle = LCL!ProgTitle
    '        If Not IsNull(LCL!ProgVer) Then Net!ProgVer = LCL!ProgVer
    '        If Not IsNull(LCL!XTitle) Then Net!XTitle = LCL!XTitle
    '        If Not IsNull(LCL!Ytitle) Then Net!Ytitle = LCL!Ytitle
    '        If Not IsNull(LCL!Notes) Then Net!Ytitle = LCL!Notes
    '        Net.Update
    '        Net.Close
    '        LCL.MoveNext
    '       '  If ExpectedProgress.Value + 1 < ExpectedProgress.Max Then ExpectedProgress.Value = ExpectedProgress.Value + 1
    '    Next i
    '    LCL.Close
    '
    '
    '    Me.txtTitle.Text = "    PLEASE WAIT....... Deleting Local TRACE Data  ............  "
    '
    '    SQLstr = "Delete from Trace"
    '    Set data = New ADODB.Recordset
    '    data.Open SQLstr, LocalTraceConn, adOpenDynamic, adLockOptimistic
    '
    '    SQLstr = "Delete from TracePoints"
    '    Set data = New ADODB.Recordset
    '    data.Open SQLstr, LocalTraceConn, adOpenDynamic, adLockOptimistic
    '
    'Me.txtTitle.Text = MessageTxt
    'If SQLAccess Then
    '        txtNet.Text = "SQL Database"
    '    ElseIf NetworkAccess Then
    '        txtNet.Text = "Network Access Database"
    '    Else
    '        txtNet.Text = "Local Access Database"
    '    End If
    'txtNet.ForeColor = &H80FF&
    'Network.BackColor = &H80FF&
    'Exit Function
    'Trap:
    '    txtNet.Text = "Local"
    '    txtNet.ForeColor = &H80FF&
    '    Network.BackColor = &H80FF&
    'End Function

    Public Sub UndoUUT(Fail As Boolean)
        Dim SQLstr As String

        If MsgBox("Are you sure you want to erase UUT" + Me.UUTCount.Text, vbOKCancel, "???") = vbCancel Then Exit Sub
        SQLstr = "Delete from TestData where JobNumber = '" + Me.cmbJob.Text & "' And SerialNumber = 'UUT" & UUTNum_Reset & "'"
        SQL.ExecuteSQLCommand(SQLstr, "NetworkData")

        If TweakMode Then
            Me.UUTMessage.Text = "  UUT TESTS  --   Testing Undisclosed Unit"
        ElseIf Not TraceChecked Then
            UUTMessage.Text = "  UUT TESTS Marker Mode  --   Load Unit #" & UUTNum + 1
        Else
            Me.UUTMessage.Text = "  UUT TESTS  --   Testing Unit #" & UUTNum
        End If

        Total1.Text = Total1.Text - 1
        Total2.Text = Total2.Text - 1
        Total3.Text = Total3.Text - 1
        Total4.Text = Total4.Text - 1
        Total5.Text = Total5.Text - 1
        If TweakMode Then
            UUTNum = 1
            EraseTest.Text = "Remove UUT" & UUTNum
            UUTNum_Reset = 1
        Else
            UUTNum = UUTNum - 1
            EraseTest.Text = "Remove UUT" & UUTNum
            UUTNum_Reset = UUTNum_Reset - 1
        End If

        UUTCount.Text = UUTNum


        'Test3


        Me.cmdStartTest.Text = "Start Test"
        Me.cmdStartTest.Enabled = True
        Me.cmdRetest.Text = "Re - Test"
        Me.cmdRetest.Enabled = False
        EraseTest.Enabled = False
        RemoveSerialEntryLog(Me.UUTCount.Text)
        If UUTNum = 0 Then
            Me.Failures1.Text = "0%"
            Me.Failures2.Text = "0%"
            Me.Failures3.Text = "0%"
            Me.Failures4.Text = "0%"
            Me.Failures5.Text = "0%"
        Else

            Me.Failures1.Text = FormatPercent(((Test1Fail / UUTNum)), 1)
            Me.Failures2.Text = FormatPercent(((Test2Fail / UUTNum)), 1)
            Me.Failures3.Text = FormatPercent(((TEST3Fail / UUTNum)), 1)
            Me.Failures4.Text = FormatPercent(((TEST4Fail / UUTNum)), 1)
            Me.Failures5.Text = FormatPercent(((TEST5Fail / UUTNum)), 1)
        End If


        ResetTests()
    End Sub

    Private Sub UUTCount_Change()
        If TweakMode Then
            UUTNum = 1
            UUTCount.Text = UUTNum
            EraseTest.Text = "Remove UUT" & UUTNum

            UUTMessage.Text = "  UUT TESTS  --   Tweak Mode. No Data Logging"
        ElseIf Not TraceChecked Then
            UUTMessage.Text = "  UUT TESTS Marker Mode  --   Load Unit #" & UUTNum + 1
            If Not IsNumeric(UUTCount.Text) Then UUTCount.Text = UUTNum
            UUTNum = UUTCount.Text
        Else
            If Not IsNumeric(UUTCount.Text) Then UUTCount.Text = UUTNum
            UUTNum = UUTCount.Text
            UUTMessage.Text = "  UUT TESTS  --   Load Unit #" & UUTCount.Text + 1
        End If


    End Sub

    Private Sub FindOffsets()
        Dim PassFail As String
        Dim RetrnVal1 As Double
        Dim RetrnVal2 As Double
        Dim Title As String

        If DontclickTheButton = True Then Exit Sub
        If Me.cmbJob.Text = "" Or Me.cmbJob.Text = " " Then
            MsgBox("Please select Job")
            Exit Sub
        End If
        FailCount = 0
        If Me.cmbJob.Text = "" Then
            MsgBox("Please Choose the Job Number first", , "Not Ready to Start")
            Exit Sub
        Else
            If Me.ckTest1.Checked Then txtOffset1.Text = 0
            If Me.ckTest2.Checked Then txtOffset2.Text = 0
            If Me.ckTest3.Checked Then txtOffset3.Text = 0
            If Me.ckTest4.Checked Then txtOffset4.Text = 0
            If Me.ckTest5.Checked Then txtOffset5.Text = 0

            Title = Me.txtTitle.Text
            RunningOffsets = True
            Me.RunStatus.ForeColor = Color.Red
            UUTCount.Text = Str(UUTNum)
            If Me.cmbJob.Text = " " Then Exit Sub

            'Insertion Loss
            If Me.ckTest1.Checked Then
                Me.txtTitle.Text = "     FINDING INSERTION LOSS Offset   SW POSITION 1      "
                Me.MutiCal.Checked = False
                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.InsertionLoss3dB(, TestID)
                If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.InsertionLossCOMB(, TestID)
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.InsertionLossDIR(, TestID)
                RetrnVal1 = IL

                Me.MutiCal.Checked = False
                SetupVNA(True, 1)
                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.InsertionLoss3dB(, TestID)
                If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.InsertionLossCOMB(, TestID)
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = Tests.InsertionLossDIR(, TestID)
                RetrnVal2 = IL

                txtOffset1.Text = TruncateDecimal((RetrnVal1 - RetrnVal2), 1)
                txtOffset_DblClick(0)
                Data1.Text = ""
            End If
            'Return Loss
            If Me.ckTest2.Checked Then
                Me.MutiCal.Checked = True
                PassFail = Tests.ReturnLoss()
                RetrnVal1 = RL
                Me.MutiCal.Checked = False
                PassFail = Tests.ReturnLoss()
                RetrnVal2 = RL
                TEST1PASS = RetrnVal1 - RetrnVal2
                txtOffset2.Text = TruncateDecimal((RetrnVal1 - RetrnVal2), 1)
                txtOffset_DblClick(1)
                Data2.Text = ""
            End If

            If SpecType <> "COMBINER/DIVIDER" Then GoTo Test2Sub
Test2SubRet:

            'AmplitudeBalance
            'Directivity
            If Me.ckTest4.Checked Then
                Me.txtTitle.Text = "     FINDING AMPLITUDE BALANCE Offset   SW POSITION 1      "
                Me.MutiCal.Checked = 1
                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.AmplitudeBalance(, 1)
                If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.AmplitudeBalanceCOMB(, 1)
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then Tests.Directivity(1, SpecType)
                RetrnVal1 = AB

                Me.MutiCal.Checked = 0
                SetupVNA(True, 1)
                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = Tests.AmplitudeBalance(, 1)
                If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.AmplitudeBalanceCOMB(, 1)
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then Tests.Directivity(1, SpecType)
                RetrnVal2 = AB

                If SpecType <> "SINGLE DIRECTIONAL COUPLER" And SpecType <> "DUAL DIRECTIONAL COUPLER" And SpecType <> "BI DIRECTIONAL COUPLER" Then
                    Test = RetrnVal1 - RetrnVal2
                    txtOffset4.Text = TruncateDecimal(RetrnVal1 - RetrnVal2, 1)
                    txtOffset_DblClick(3)
                    If SpecAB_TF Then
                        Data4L.Text = ""
                        Data4H.Text = ""
                    Else
                        Data4.Text = ""
                    End If
                End If
            End If

            'PhaseBalance
            'CoupledFlatness
            If Me.ckTest5.Checked Then
                ActiveTitle = "     FINDING PHASE BALANCE Offset   SW POSITION 1      "
                Me.MutiCal.Checked = True
                If SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN" Then PassFail = Tests.PhaseBalance(SpecType)
                If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.PhaseBalanceCOMB(SpecType)
                RetrnVal1 = PB
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then Tests.CoupledFlatness(1, SpecType)
                RetrnVal1 = COuP

                Me.MutiCal.Checked = False
                SetupVNA(True, 1)
                If SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN" Then PassFail = Tests.PhaseBalance(SpecType)
                If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.PhaseBalanceCOMB(SpecType, RetrnVal2)
                RetrnVal1 = PB
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then Tests.CoupledFlatness(1, SpecType)
                RetrnVal1 = COuP
                If SpecType <> "SINGLE DIRECTIONAL COUPLER" And SpecType <> "DUAL DIRECTIONAL COUPLER" Then
                    Test = RetrnVal1 - RetrnVal2
                    txtOffset3.Text = TruncateDecimal((RetrnVal1 - RetrnVal2), 1)

                    txtOffset_DblClick(4)
                    Data5.Text = ""
                End If
            End If
            If SpecType <> "COMBINER/DIVIDER" Then GoTo TestComplete
Test2Sub:

            'Isolation
            'Coupling
            If Me.ckTest3.Checked Then
                ActiveTitle = "     FINDING ISOLATION Offset   SW POSITION 1      "
                Me.MutiCal.Checked = True
                If SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN" Then PassFail = Tests.Isolation(RetrnVal1)
                If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.IsolationCOMB(RetrnVal1)
                If ISO_TF Then
                    RetrnVal1 = ISoL
                    RetrnVal2 = ISoH
                Else
                    RetrnVal1 = ISo
                End If

                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then Tests.Coupling(RetrnVal1, 1, SpecType)
                RetrnVal1 = COuP

                Me.MutiCal.Checked = False
                SetupVNA(True, 1)
                If SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN" Then PassFail = Tests.Isolation(RetrnVal2)
                If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = Tests.IsolationCOMB(RetrnVal2)
                RetrnVal2 = ISo
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then Tests.Coupling(RetrnVal2, 1, SpecType)
                RetrnVal2 = COuP

                If SpecType <> "SINGLE DIRECTIONAL COUPLER" And SpecType <> "DUAL DIRECTIONAL COUPLER" Then
                    txtOffset3.Text = TruncateDecimal((RetrnVal2 - RetrnVal1), 1)
                    txtOffset_DblClick(2)
                    If ISO_TF Then
                        Data3L.Text = ""
                        Data3H.Text = ""
                    Else
                        Data3.Text = ""
                    End If

                End If

            End If
            If SpecType <> "COMBINER/DIVIDER" Then GoTo Test2SubRet
TestComplete:  ' For everything except Directional Couplers

            'Directonal Couplers reverse direction
            If Not TweakMode And (SpecType = "DUAL DIRECTIONAL COUPLER" Or (SpecType = "BI DIRECTIONAL COUPLER" And Me.ckTest4.Text)) Then
                MsgBox("Please turn the Directional Coupler in the Reverse direction")
            End If

            'Reverse Coupling
            If SpecType = "DUAL DIRECTIONAL COUPLER" Then
                If Me.ckTest3.Checked Then
                    Me.MutiCal.Checked = True
                    Tests.Coupling(RetrnVal1, 1, SpecType)

                    Me.MutiCal.Checked = False
                    SetupVNA(True, 1)
                    Tests.Coupling(RetrnVal2, 1, SpecType)

                    txtOffset3.Text = TruncateDecimal((RetrnVal2 - RetrnVal1), 1)
                    txtOffset_DblClick(2)
                    Data3.Text = ""
                End If
            End If

            If SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then
                'Reverse Directivity
                If Me.ckTest4.Checked Then
                    Me.MutiCal.Checked = True
                    Tests.Directivity(2, SpecType)
                    RetrnVal1 = DIR

                    Me.MutiCal.Checked = False
                    SetupVNA(True, 1)
                    Tests.Directivity(2, SpecType)
                    RetrnVal2 = DIR

                    Test = RetrnVal1 - RetrnVal2
                    txtOffset4.Text = TruncateDecimal(RetrnVal1 - RetrnVal2, 1)
                    txtOffset_DblClick(3)
                    If SpecAB_TF Then
                        Data4L.Text = ""
                        Data4H.Text = ""
                    Else
                        Data4.Text = ""
                    End If
                End If
            End If
            If SpecType = "DUAL DIRECTIONAL COUPLER" Then
                'Reverse CoupledFlatness
                If Me.ckTest5.Checked Then
                    Me.MutiCal.Checked = True
                    Tests.CoupledFlatness(1, SpecType)
                    RetrnVal1 = COuP

                    Me.MutiCal.Checked = False
                    SetupVNA(True, 1)
                    Tests.CoupledFlatness(1, SpecType)
                    RetrnVal2 = CF

                    Test = RetrnVal1 - RetrnVal2
                    txtOffset5.Text = TruncateDecimal((RetrnVal1 - RetrnVal2), 1)
                    txtOffset_DblClick(4)
                    Data5.Text = ""
                End If
            End If
        End If
        cmdStartTest.Text = "Start   Test"
        RunningOffsets = False
        ActiveTitle = Title
        Me.GetOffsets.Checked = False

    End Sub



    Private Sub mnuScreenShot_Click()
        'Dim Plotstr As String
        'Dim hFile As Long
        'Dim FileNameWithPath As String
        'Dim semi_len As Long

        'Plotstr = "qlkwkjwoi;lkwp;o;eoie;ihe;i"

        '        Plotstr = ScanGPIB.DeviceQuery("OUTPPLOT;")
        '        FileNameWithPath = "C:\ATS\Plot.bmp"
        '        Open Trim(FileNameWithPath) ' For Output As #1

        '        semi_len = InStr(Plotstr, ";")
        '        Plotstr = Mid(Plotstr, semi_len, Len(Plotstr) - semi_len)
        'Print (#1, Plotstr

        'Close #1



        'hFile = CreateFile(FileNameWithPath, GENERIC_READ, FILE_SHARE_READ, _
        '        ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0&)
        '
        'CloseHandle (hFile)

        'ScreenShot.Show(vbModeless, Me)


    End Sub
    Private Sub Simulation_Changed(sender As Object, e As EventArgs) Handles Simulation.CheckedChanged
        If Simulation.Checked Then
            Me.RealData.Visible = True
            Me.Pass.Visible = True
            Me.Fail.Visible = True
            StatusLog.Items.Add("Data Simulation Mode: ON")
            DebugChecked = True
        Else
            StatusLog.Items.Add("Data Simulation Mode: OFF")
            Me.RealData.Visible = False
            Me.Pass.Visible = False
            Me.Fail.Visible = False
            DebugChecked = False
        End If
        Debug = Simulation.Checked

    End Sub
    Private Sub LoadSpecs()
        Try
            SpecStartFreq = Nothing
            SpecStopFreq = Nothing
            SpecIL = Nothing
            SpecRL = Nothing
            SpecISO = Nothing
            SpecISOL = Nothing
            SpecISOH = Nothing
            SpecAB = Nothing
            SpecPB = Nothing
            SpecCOUP = Nothing
            SpecCOUPPM = Nothing
            SpecDIRECT = Nothing
            SpecCOUPFLAT = Nothing
            SpecPorts = Nothing

            SpecIL = CDbl(SQL.GetSpecification("InsertionLoss"))
            SpecILL = CDbl(SQL.GetSpecification("InsertionLoss"))
            SpecILH = CDbl(SQL.GetSpecification("IL_ex"))
            SpecRL = CDbl(SQL.GetSpecification("VSWR"))
            SpecISO = CDbl(SQL.GetSpecification("Isolation"))
            SpecISOL = CDbl(SQL.GetSpecification("Isolation"))
            SpecISOH = CDbl(SQL.GetSpecification("Isolation2"))
            SpecCuttoffFreq = CDbl(SQL.GetSpecification("CutOffFreqMHz"))
            SpecCOUP = CDbl(SQL.GetSpecification("Coupling"))
            SpecCOUPPM = CDbl(SQL.GetSpecification("CouplingPM"))
            SpecDIRECT = CDbl(SQL.GetSpecification("Directivity"))
            SpecCOUPFLAT = CDbl(SQL.GetSpecification("CoupledFlatness"))
            SpecAB = CDbl(SQL.GetSpecification("AmplitudeBalance"))
            SpecPB = CDbl(SQL.GetSpecification("PhaseBalance"))
            SpecPorts = CDbl(SQL.GetSpecification("OutputPortNumber"))
            SpecStartFreq = CDbl(SQL.GetSpecification("StartFreqMHz"))
            SpecStopFreq = CDbl(SQL.GetSpecification("StopFreqMHz"))
            If SQL.GetBypass() = 1 Then
                Master_bypass = True
            Else
                Master_bypass = False
            End If
            TestFailMax = Quantity * GetFailPercent() / 100
            GlobalFailMax = Quantity * GetFailPercent() / 100


        Catch ex As Exception
        End Try
    End Sub

    Private Sub txtNet_CheckedChanged(sender As Object, e As EventArgs) Handles txtNet.CheckedChanged
        NetworkChecked()
    End Sub

    Private Sub NetworkChecked()
        If txtNet.Checked Then
            If SQLVerified Then
                SQLAccess = txtNet.Checked
            Else
                MsgBox("The SQL Server was not available at Start Up", , "Can't do it!!!")
                SQLAccess = False
                txtNet.Checked = False
            End If

            If SQLAccess Then
                txtNet.Text = " SQL Database"
                txtNet.ForeColor = Color.LawnGreen
            Else
                If NetworkAccess Then
                    txtNet.Text = "Nework Access Database"
                    txtNet.ForeColor = Color.DarkOrange
                Else
                    txtNet.Text = "Local Access Database"
                    txtNet.ForeColor = Color.DarkOrange
                End If
            End If
        Else
            If NetworkAccess Then
                txtNet.Text = "Nework Access Database"
                txtNet.ForeColor = Color.DarkOrange
            Else
                txtNet.Text = "Local Access Database"
                txtNet.ForeColor = Color.DarkOrange
            End If
        End If
    End Sub


    Private Sub ScanGPIBToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ScanGPIBToolStripMenuItem.Click

        'Dim CalDone As Boolean
        StatusLog.Items.Add("Calibration Started" & "" & DateTime.Now.ToString)
        Tests.CalibrateVNA()
        'If CalDone Then CalDueDate.Show(vbModal, Me)
        ClearStatusLog()
    End Sub



    Private Sub MutiCal_CheckedChanged(sender As Object, e As EventArgs) Handles MutiCal.CheckedChanged
        StatusLog.Items.Add("Multical Mode:" & "" & DateTime.Now.ToString)
        MutiCalChecked = Me.MutiCal.Checked
    End Sub

    Private Sub RFSwitch_CheckedChanged(sender As Object, e As EventArgs) Handles RFSwitch.CheckedChanged
        SwitchedChecked = Me.RFSwitch.Checked
    End Sub


    Private Sub GetTrace_CheckedChanged(sender As Object, e As EventArgs) Handles GetTrace.CheckedChanged
        If Not TraceChecked And Me.GetTrace.CheckState = CheckState.Checked And UUTNum > 5 Then
            BypassUnchecked = True
        ElseIf Me.GetTrace.CheckState = CheckState.Unchecked Then
            BypassUnchecked = False
        End If
        TraceChecked = Me.GetTrace.Checked
    End Sub


    Private Sub ExtraAverage_CheckedChanged(sender As Object, e As EventArgs) Handles ExtraAverage.CheckedChanged
        EAveraging = Me.ExtraAverage.Checked
        If Not Simulation.Checked And Not Connected Then
            ScanGPIB.connect("GPIB0::16::INSTR", GetTimeout())
            Connected = True
        End If
        Averages = AvgS.Text
        If Not Debug Then
            If ExtraAverage.Checked Then
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite("SENS:AVER:COUNT " & CStr(AvgS.Text))
                    ScanGPIB.BusWrite("SENS:AVER ON")
                    ScanGPIB.BusWrite("SENS2:AVER:COUNT " & CStr(AvgS.Text))
                    ScanGPIB.BusWrite("SENS2:AVER ON")
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("SENS:AVER:COUNT " & CStr(AvgS.Text))
                    ScanGPIB.BusWrite("SENS:AVER ON")
                    ScanGPIB.BusWrite("SENS2:AVER:COUNT " & CStr(AvgS.Text))
                    ScanGPIB.BusWrite("SENS2:AVER ON")
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN1;")
                    ScanGPIB.BusWrite("AVERFACT " & CStr(AvgS.Text))
                    ScanGPIB.BusWrite("AVERO ON")
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("AVERFACT " & CStr(AvgS.Text))
                    ScanGPIB.BusWrite("AVERO ON")
                End If
            Else
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite("SENS:AVER OFF")
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("SENS:AVER OFF")
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN1;")
                    ScanGPIB.BusWrite("AVERO OFF")
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("AVERO OFF")
                End If
            End If
        End If
    End Sub

    Private Sub TestDataToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim TD As New TestDataView
        Me.Hide()
        StatusLog.Items.Add("Opening Test Data:" & "" & DateTime.Now.ToString)
        TD.StartPosition = FormStartPosition.Manual
        TD.Location = New Point(globals.XLocation, globals.YLocation)
        TD.ShowDialog()
        Me.Show()
    End Sub

    Private Sub SpecificationsToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        Dim spec As New SpecDataView
        Me.Hide()
        StatusLog.Items.Add("Opening Specs:" & "" & DateTime.Now.ToString)
        spec.StartPosition = FormStartPosition.Manual
        spec.Location = New Point(globals.XLocation, globals.YLocation)
        spec.ShowDialog()
        Me.Show()
    End Sub

    Private Sub PortConfigurationsToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim CONFIG As New PortConfigView
        Me.Hide()
        StatusLog.Items.Add("Opening Port Config's:" & "" & DateTime.Now.ToString)
        CONFIG.StartPosition = FormStartPosition.Manual
        CONFIG.Location = New Point(globals.XLocation, globals.YLocation)
        CONFIG.ShowDialog()
        Me.Show()
    End Sub

    Private Sub TestEquipmentToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim DEV As New DeviceView
        Me.Hide()
        StatusLog.Items.Add("Opening Devices:" & "" & DateTime.Now.ToString)
        DEV.StartPosition = FormStartPosition.Manual
        DEV.Location = New Point(globals.XLocation, globals.YLocation)
        DEV.ShowDialog()
        Me.Show()
    End Sub

    Private Sub WorkstationsToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim DEV As New WorkStationView
        Me.Hide()
        StatusLog.Items.Add("Opening WorkStations:" & "" & DateTime.Now.ToString)
        DEV.StartPosition = FormStartPosition.Manual
        DEV.Location = New Point(globals.XLocation, globals.YLocation)
        DEV.ShowDialog()
        Me.Show()
    End Sub

    Private Sub SpecificationsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpecificationsToolStripMenuItem.Click
        Dim SPEC As New frmSpecifications
        Me.Hide()
        StatusLog.Items.Add("Opening Specs:" & "" & DateTime.Now.ToString)
        SPEC.StartPosition = FormStartPosition.Manual
        SPEC.Location = New Point(globals.XLocation, globals.YLocation)
        SPEC.ShowDialog()
        LoadSpecs()
        Me.Show()
    End Sub


    Private Sub RealData_CheckedChanged(sender As Object, e As EventArgs) Handles RealData.CheckedChanged
        If RealData.Checked Then
            DBDataChecked = True
        Else
            DBDataChecked = False
        End If
    End Sub

    Private Sub Pass_CheckedChanged(sender As Object, e As EventArgs) Handles Pass.CheckedChanged
        If Pass.Checked Then
            PassChecked = True
        Else
            PassChecked = False
        End If
    End Sub

    Private Sub Fail_CheckedChanged(sender As Object, e As EventArgs) Handles Fail.CheckedChanged
        If Fail.Checked Then
            FailChecked = True
        Else
            FailChecked = False
        End If
    End Sub


    Private Sub TraceSearchToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim TRD As New TraceView
        Me.Hide()
        StatusLog.Items.Add("Opening Trace Data:" & "" & DateTime.Now.ToString)
        TRD.StartPosition = FormStartPosition.Manual
        TRD.Location = New Point(globals.XLocation, globals.YLocation)
        TRD.ShowDialog()
        Me.Show()
    End Sub

    Private Sub OperatorEffeciencyToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Dim EFF As New EffeciencyView
        Me.Hide()
        StatusLog.Items.Add("Opening Operator Effeciency:" & "" & DateTime.Now.ToString)
        EFF.StartPosition = FormStartPosition.Manual
        EFF.Location = New Point(globals.XLocation, globals.YLocation)
        EFF.ShowDialog()
        Me.Show()

    End Sub
    Public Function ReportServer(Status As String, Optional value As Integer = 0, Optional first As Boolean = False) As String
        Dim SQLstr As String
        Dim ReportType As String = SpecType
        Dim history As String = GetreportStatus()
        ReportServer = history

        SQLstr = "SELECT * from ReportQueue where JobNumber = '" & Job & "' And WorkStation = '" & GetComputerName() & "'"
        If SQL.CheckforRow(SQLstr, "ReportQueue") = 0 Then
            SQLstr = "Insert Into ReportQueue (ReportName, ReportType,ReportStatus,JobNumber,WorkStation,PartNumber,Operator,ActiveDate,MaxValue) values ('" & SpecType & "','" & ReportType & "','" & Status & "','" & Job & "','" & GetComputerName() & "','" & Part & "','" & User & "','" & DateTime.Now.ToString & "','" & Quantity & "')"
            SQL.ExecuteSQLCommand(SQLstr, "NetworkData")
        Else
            SQLstr = "UPDATE ReportQueue Set ReportStatus = " & Status & " where JobNumber = '" & Job & "' And WorkStation = '" & GetComputerName() & "'"
            SQL.ExecuteSQLCommand(SQLstr, "Effeciency")
            SQLstr = "UPDATE ReportQueue Set Value = " & value & " where JobNumber = '" & Job & "' And WorkStation = '" & GetComputerName() & "'"
            SQL.ExecuteSQLCommand(SQLstr, "Effeciency")
            SQLstr = "UPDATE ReportQueue Set ReportStatus = '" & Status & "' where JobNumber = '" & Job & "' And WorkStation = '" & GetComputerName() & "'"
            SQL.ExecuteSQLCommand(SQLstr, "Effeciency")
        End If


    End Function
    Public Sub ExcelData()

        Dim SQLstr As String
        Dim Test As String
        Dim Directional As Boolean
        Dim Row As Long
        Dim TopFolder As String
        Dim SubFolder As String
        Dim Temp As String

        Try
            Directional = False
            Row = 6
            SpecType = SQL.GetSpecification("SpecType")
            If SpecType.Contains("DIRECTIONAL COUPLER") Then Directional = True

            'SQLstr = "Select * from TestData where JobNumber = '" & Me.cmbJob.Text & "' and Workstation = '" & WorkStation & "'"
            SQLstr = "Select * from TestData where JobNumber = '" & Me.cmbJob.Text & "'"

            Test = SQL.CheckforRow(SQLstr, "NetworkData")
            If Test = 0 Then
                MsgBox("Sorry no data on record")
                Exit Sub
            End If
            Temp = txtTitle.Text
            txtTitle.Text = "Saving Test Report..... Please Wait"
            If WorkStation = "Autom" Then
                ExcelReports.StartupReport("C:\ATS\Client_IPP\IPP_Coupler Automation PROJECT\Excel Templates\", "TestData.xlsm")
            Else
                ExcelReports.StartupReport(ExcelTemplatePath, "TestData1.xls")
            End If

            ExcelReports.DataToCells("G2", Job)
            ExcelReports.DataToCells("G3", Part)
            ExcelReports.DataToCells("G4", SpecType)
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                If Directional Then
                    ExcelReports.DataToCells("E5", "Coupling")
                    ExcelReports.DataToCells("F5", "Directivity")
                    ExcelReports.DataToCells("G5", "Coupling Flatness")
                    While Not dr.Read = Nothing
                        If Not IsDBNull(dr.Item(4)) Then ExcelReports.DataToCells("A" & Row, dr.Item(4))
                        If Not IsDBNull(dr.Item(5)) Then ExcelReports.DataToCells("B" & Row, dr.Item(5))
                        If Not IsDBNull(dr.Item(6)) Then ExcelReports.DataToCells("C" & Row, dr.Item(6))
                        If Not IsDBNull(dr.Item(7)) Then ExcelReports.DataToCells("D" & Row, Format(dr.Item(7), "0.00"))
                        If Not IsDBNull(dr.Item(8)) Then ExcelReports.DataToCells("E" & Row, Format(dr.Item(8), "0.00"))
                        If Not IsDBNull(dr.Item(10)) Then ExcelReports.DataToCells("F" & Row, Format(dr.Item(10), "0.00"))
                        If Not IsDBNull(dr.Item(12)) Then ExcelReports.DataToCells("G" & Row, Format(dr.Item(12), "0.00"))
                        If IsDBNull(dr.Item(15)) Then
                            ExcelReports.DataToCells("H" & Row, ArtworkRevision)
                        Else
                            ExcelReports.DataToCells("H" & Row, dr.Item(15))
                        End If

                        Row = Row + 1
                    End While
                Else

                    If Not SpecType.Contains("BALUN") Then
                        ExcelReports.DataToCells("E5", "Isolation")
                    Else
                        ExcelReports.DataToCells("E5", "No Test")
                    End If

                    ExcelReports.DataToCells("F5", "Amplitude Balance")
                    ExcelReports.DataToCells("G5", "Phase Balance")
                    While Not dr.Read = Nothing
                        If Not IsDBNull(dr.Item(4)) Then ExcelReports.DataToCells("A" & Row, dr.Item(4))
                        If Not IsDBNull(dr.Item(5)) Then ExcelReports.DataToCells("B" & Row, dr.Item(5))
                        If IL_TF Then
                            If Not IsDBNull(dr.Item(6)) And Not IsDBNull(dr.Item(24)) Then ExcelReports.DataToCells("C" & Row, dr.Item(6) & " / " & dr.Item(24))
                        Else
                            If Not IsDBNull(dr.Item(6)) Then ExcelReports.DataToCells("C" & Row, dr.Item(6))
                        End If
                        If Not IsDBNull(dr.Item(7)) Then ExcelReports.DataToCells("D" & Row, Format(dr.Item(7), "0.00"))
                        If IL_TF Then
                            ExcelReports.DataToCells("E" & Row, "N/A")
                            ExcelReports.DataToCells("F" & Row, "N/A")
                            ExcelReports.DataToCells("G" & Row, "N/A")
                            ExcelReports.DataToCells("H" & Row, "N/A")
                        Else
                            If Not SpecType.Contains("BALUN") Then If Not IsDBNull(dr.Item(9)) Then ExcelReports.DataToCells("E" & Row, Format(dr.Item(9), "0.00"))
                            If SpecAB_TF Then
                                If Not IsDBNull(dr.Item(16)) And Not IsDBNull(dr.Item(17)) Then ExcelReports.DataToCells("F" & Row, dr.Item(16) & " / " & dr.Item(17))
                            Else
                                If Not IsDBNull(dr.Item(11)) Then ExcelReports.DataToCells("F" & Row, dr.Item(11))
                            End If
                            If Not IsDBNull(dr.Item(13)) Then ExcelReports.DataToCells("G" & Row, Format(dr.Item(13), "0.00"))
                            If IsDBNull(dr.Item(15)) Then
                                ExcelReports.DataToCells("H" & Row, ArtworkRevision)
                            Else
                                ExcelReports.DataToCells("H" & Row, dr.Item(15))
                            End If
                        End If
                        Row = Row + 1
                    End While
                    ats.Close()
                End If
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                If Directional Then
                    ExcelReports.DataToCells("E5", "Coupling")
                    ExcelReports.DataToCells("E6", "Directivity")
                    ExcelReports.DataToCells("E7", "Coupling Balance")
                    While Not drLocal.Read = Nothing
                        If Not IsDBNull(drLocal.Item(4)) Then ExcelReports.DataToCells("A" & Row, drLocal.Item(4))
                        If Not IsDBNull(drLocal.Item(5)) Then ExcelReports.DataToCells("B" & Row, drLocal.Item(5))
                        If Not IsDBNull(drLocal.Item(6)) Then ExcelReports.DataToCells("C" & Row, drLocal.Item(6))
                        If Not IsDBNull(drLocal.Item(7)) Then ExcelReports.DataToCells("D" & Row, drLocal.Item(7))
                        If Not IsDBNull(drLocal.Item(8)) Then ExcelReports.DataToCells("E" & Row, drLocal.Item(8))
                        If Not IsDBNull(drLocal.Item(10)) Then ExcelReports.DataToCells("F" & Row, drLocal.Item(10))
                        If Not IsDBNull(drLocal.Item(12)) Then ExcelReports.DataToCells("G" & Row, drLocal.Item(12))
                        If IsDBNull(drLocal.Item(15)) Then
                            ExcelReports.DataToCells("H" & Row, ArtworkRevision)
                        Else
                            ExcelReports.DataToCells("H" & Row, drLocal.Item(15))
                        End If
                        Row = Row + 1
                    End While
                Else
                    If Not SpecType.Contains("BALUN") Then
                        ExcelReports.DataToCells("E5", "Iso")
                    Else
                        ExcelReports.DataToCells("E5", "No Test")
                    End If
                    ExcelReports.DataToCells("F5", "Amplitude Balance")
                    ExcelReports.DataToCells("G5", "Phase Balance")
                    While Not drLocal.Read = Nothing
                        If Not IsDBNull(drLocal.Item(4)) Then ExcelReports.DataToCells("A" & Row, drLocal.Item(4))
                        If Not IsDBNull(drLocal.Item(5)) Then ExcelReports.DataToCells("B" & Row, drLocal.Item(5))
                        If Not IsDBNull(drLocal.Item(6)) Then ExcelReports.DataToCells("C" & Row, drLocal.Item(6))
                        If Not IsDBNull(drLocal.Item(7)) Then ExcelReports.DataToCells("D" & Row, drLocal.Item(7))
                        If Not SpecType.Contains("BALUN") Then If Not IsDBNull(drLocal.Item(9)) Then ExcelReports.DataToCells("E" & Row, drLocal.Item(9))
                        If SpecAB_TF Then
                            If Not IsDBNull(drLocal.Item(16)) And Not IsDBNull(drLocal.Item(17)) Then ExcelReports.DataToCells("F" & Row, drLocal.Item(16) & "/" & drLocal.Item(17))
                        Else
                            If Not IsDBNull(drLocal.Item(11)) Then ExcelReports.DataToCells("F" & Row, drLocal.Item(11))
                        End If
                        If Not IsDBNull(drLocal.Item(13)) Then ExcelReports.DataToCells("G" & Row, drLocal.Item(13))
                        If IsDBNull(drLocal.Item(15)) Then
                            ExcelReports.DataToCells("H" & Row, ArtworkRevision)
                        Else
                            ExcelReports.DataToCells("H" & Row, drLocal.Item(15))
                        End If
                        Row = Row + 1
                    End While
                    atsLocal.Close()
                End If
            End If
            TraceDataReport()
            TopFolder = "90_Degree\"
            If SpecType = "90 DEGREE COUPLER" Then TopFolder = "90_Degree\"
            If SpecType.Contains("BALUN") Then TopFolder = "Balun\"
            If SpecType = "90 DEGREE COUPLER SMD" Then TopFolder = "90_Degree_SMD\"
            If InStr(SpecType, "DIRECTIONAL COUPLER") Then TopFolder = "Directional_Couplers\"
            If InStr(SpecType, "DIRECTIONAL COUPLER SMD") Then TopFolder = "Directional_Couplers_SMD\"
            If SpecType.Contains("COMBINER/DIVIDER") Then TopFolder = "Combiner-Divider\"
            If SpecType = "COMBINER/DIVIDER SMD" Then TopFolder = "Combiner-Divider_SMD\"

            If NetworkAccess Then
                'Network Save
                If Not System.IO.Directory.Exists(TestDataPath & TopFolder) Then System.IO.Directory.CreateDirectory(TestDataPath & TopFolder)
                SubFolder = TestDataPath & TopFolder & Me.cmbPart.Text & "-" & Trim(Me.cmbJob.Text) & "\"
                If Not System.IO.Directory.Exists(SubFolder) Then System.IO.Directory.CreateDirectory((SubFolder))
            Else
                'Local Save
                If Not System.IO.Directory.Exists("C:\ATE Data\" & TopFolder) Then System.IO.Directory.CreateDirectory("C:\ATE Data\" & TopFolder)
                SubFolder = "C:\ATE Data\" & TopFolder & Trim(Me.cmbJob.Text) & "-" & Trim(Me.cmbJob.Text) & "\"

            End If
            If Not System.IO.Directory.Exists(SubFolder) Then System.IO.Directory.CreateDirectory((SubFolder))

            ExcelReports.SaveAs(SubFolder & "TestData " & Me.cmbJob.Text & ".xls")
            ' ExcelReports.EndReport()
            txtTitle.Text = Temp
        Catch
            ' MsgBox("An Error has Occured In The TestData Report" & vbCr & "Report This Error To AutomatedTestSolutions@Gmail.com" & vbCr & "Error Details :-" & vbCr & "Error Number : " & Err.Number & vbCr & "Error Description : " & Err.Description, vbCritical, "F")
        End Try

    End Sub

    Public Sub TraceDataReport()
        Try
            If SpecType.Contains("90 DEGREE COUPLER") Or SpecType.Contains("BALUN") Or SpecType.Contains("TRANSFORMER") Then

                For x = 0 To 4
                    SerialNumber = "UUT" & x + 1

                    If Not IL_XArray(0, 0) = 0 Then
                        For y = 0 To 200
                            ReDim Preserve XArray(y)
                            ReDim Preserve YArray(y)
                            ReDim Preserve YArray1(y)
                            XArray(y) = IL_XArray(x, y)
                            YArray(y) = IL1_YArray(x, y)
                            YArray1(y) = IL2_YArray(x, y)
                        Next
                    Else
                        Title = "Insertion Loss J4"
                        If Not NetworkAccess Then
                            TraceID = 4263
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID = 0 Then
                            GoTo RL
                        ElseIf TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray1(YArray1.Count)
                        YArray1 = YArray

                        Title = "Insertion Loss J3"
                        If Not NetworkAccess Then
                            TraceID = 4262
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray(YArray.Count)
                    End If
                    ExcelReports.LoadChart1(SerialNumber, "Insertion Loss/Amplitude balance")
RL:
                    ' Return Loss
                    If Not RL_XArray(0, 0) = 0 Then
                        For y = 0 To 200
                            ReDim Preserve XArray(y)
                            ReDim Preserve YArray(y)
                            XArray(y) = RL_XArray(x, y)
                            YArray(y) = RL_YArray(x, y)
                        Next
                    Else
                        Title = "Return Loss"
                        If Not NetworkAccess Then
                            TraceID = 4264
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID = 0 Then
                            GoTo ISO
                        ElseIf TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                    End If
                    ExcelReports.LoadChart2(SerialNumber, Title)

ISO:
                    ' Isolation
                    If Not ISO_XArray(0, 0) = 0 Then
                        For y = 0 To 200
                            ReDim Preserve XArray(y)
                            ReDim Preserve YArray(y)
                            XArray(y) = ISO_XArray(x, y)
                            YArray(y) = ISO_YArray(x, y)
                        Next
                    Else
                        Title = "Isolation"
                        If Not NetworkAccess Then
                            TraceID = 4267
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If

                        If TraceID = 0 Then
                            GoTo PB
                        ElseIf TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                    End If
                    ExcelReports.LoadChart3(SerialNumber, Title)

PB:
                    ' Phase Balance
                    If Not PB_XArray(0, 0) = 0 Then
                        For y = 0 To 200
                            ReDim Preserve XArray(y)
                            ReDim Preserve YArray(y)
                            ReDim Preserve YArray1(y)
                            XArray(y) = PB_XArray(x, y)
                            YArray(y) = PB1_YArray(x, y)
                            YArray1(y) = PB2_YArray(x, y)
                        Next
                    Else
                        Title = "Phase Balance J4"
                        If Not NetworkAccess Then
                            TraceID = 4266
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray1(YArray.Count)
                        YArray1 = YArray
                        Title = "Phase Balance J3"
                        If Not NetworkAccess Then
                            TraceID = 4263
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                    End If
                    ExcelReports.LoadChart4(SerialNumber, "Phase Balance")
Skip1:
                Next
            End If

            If SpecType.Contains("COMBINER/DIVIDER") Then
                For x = 0 To 4
                    SerialNumber = "UUT" & x + 1

                    'Insertion Loss
                    If Not IL_XArray(0, 0) = 0 Then
                        For y = 0 To 200
                            ReDim Preserve XArray(y)
                            ReDim Preserve YArray(y)
                            XArray(y) = IL_XArray(x, y)
                            YArray(y) = IL1_YArray(x, y)
                        Next
                    Else
                        Title = "Insertion Loss"
                        If Not NetworkAccess Then
                            TraceID = 4263
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID = 0 Then
                            GoTo RL2
                        ElseIf TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray1(YArray1.Count)
                        YArray1 = YArray
                    End If
                    ExcelReports.LoadChart1(SerialNumber, "Insertion Loss/Amplitude balance")
RL2:
                    'Return Loss
                    If Not RL_XArray(0, 0) = 0 Then
                        For y = 0 To 200
                            ReDim Preserve XArray(y)
                            ReDim Preserve YArray(y)
                            XArray(y) = RL_XArray(x, y)
                            YArray(y) = RL_YArray(x, y)
                        Next
                    Else
                        Title = "Return Loss"
                        If Not NetworkAccess Then
                            TraceID = 4264
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID = 0 Then
                            GoTo ISO2
                        ElseIf TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray1(YArray1.Count)
                        ExcelReports.LoadChart2(SerialNumber, Title)
                    End If
ISO2:
                    'Isolation
                    If Not ISO_XArray(0, 0) = 0 Then
                        For y = 0 To 200
                            ReDim Preserve XArray(y)
                            ReDim Preserve YArray(y)
                            XArray(y) = ISO_XArray(x, y)
                            YArray(y) = ISO_YArray(x, y)
                        Next
                    Else
                        Title = "Isolation"
                        If Not NetworkAccess Then
                            TraceID = 4265
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID = 0 Then
                            GoTo PB2
                        ElseIf TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray1(YArray1.Count)
                    End If
                    ExcelReports.LoadChart3(SerialNumber, Title)
PB2:
                    ' Phase Balance
                    If Not PB_XArray(0, 0) = 0 Then
                        For y = 0 To 200
                            ReDim Preserve XArray(y)
                            ReDim Preserve YArray(y)
                            ReDim Preserve YArray1(y)
                            XArray(y) = PB_XArray(x, y)
                            YArray(y) = PB1_YArray(x, y)
                            YArray1(y) = PB2_YArray(x, y)
                        Next
                    Else
                        Title = "Phase Balance J4"
                        If Not NetworkAccess Then
                            TraceID = 4266
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID = 0 Then
                            GoTo Skip2
                        ElseIf TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray1(YArray1.Count)
                        YArray1 = YArray
                        Title = "Phase Balance J3"
                        If Not NetworkAccess Then
                            TraceID = 4267
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray1(YArray1.Count)
                    End If
                    ExcelReports.LoadChart4(SerialNumber, "Phase Balance")
Skip2:
                Next
            End If

            If SpecType.Contains("DIRECTIONAL COUPLER") Then
                For x = 1 To 5
                    SerialNumber = "UUT" & x
                    'Insertion Loss
                    If Not IL_XArray(0, 0) = 0 Then
                        For y = 0 To 200
                            ReDim Preserve XArray(y)
                            ReDim Preserve YArray1(y)
                            XArray(y) = IL_XArray(x, y)
                            YArray1(y) = IL1_YArray(x, y)
                        Next
                    Else
                        Title = "Insertion Loss J3"
                        If Not NetworkAccess Then
                            TraceID = 4263
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID = 0 Then
                            GoTo RL3
                        ElseIf TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray1(YArray1.Count)
                        YArray1 = YArray
                    End If
                    ExcelReports.LoadChart1(SerialNumber, "Insertion Loss")
RL3:
                    'Return Loss
                    If Not RL_XArray(0, 0) = 0 Then
                        For y = 0 To 200
                            ReDim Preserve XArray(y)
                            ReDim Preserve YArray(y)
                            ReDim Preserve YArray1(y)
                            XArray(y) = RL_XArray(x, y)
                            YArray(y) = RL_YArray(x, y)
                        Next
                    Else
                        Title = "Return Loss"
                        If Not NetworkAccess Then
                            TraceID = 4264
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID = 0 Then
                            GoTo PB3
                        ElseIf TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray1(YArray1.Count)
                    End If
                    ExcelReports.LoadChart2(SerialNumber, Title)
PB3:
                    ' Phase Balance
                    If Not COUP_XArray(0, 0) = 0 Then
                        For y = 0 To 200
                            ReDim Preserve XArray(y)
                            ReDim Preserve YArray1(y)
                            ReDim Preserve YArray2(y)
                            XArray(y) = COUP_XArray(x, y)
                            YArray1(y) = COUP1_YArray(x, y)
                            YArray2(y) = COUP2_YArray(x, y)
                        Next
                    Else
                        Title = "Coupling J4"
                        If Debug Then
                            TraceID = 4263
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID = 0 Then
                            GoTo Skip3
                        ElseIf TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray2(YArray2.Count)
                        YArray1 = YArray
                        Title = "Coupling J3"
                        If Not NetworkAccess Then
                            TraceID = 4263
                        Else
                            TraceID = GetTraceIDByTitle(Title, SerialNumber, Me.cmbJob.Text, WorkStation)
                        End If
                        If TraceID > 171666 Then
                            GetTracePoints2(TraceID)
                        Else
                            GetTracePoints(TraceID)
                        End If
                        ReDim Preserve YArray1(YArray2.Count)
                        YArray2 = YArray
                    End If
                    For z = 0 To 200
                        ReDim Preserve YArray(z)
                        YArray(z) = Math.Abs(YArray2(z) - YArray1(z))
                    Next
                    ExcelReports.LoadChart3(SerialNumber, Title)
                    ExcelReports.LoadChart4(SerialNumber, "Coupling Balance")
Skip3:
                Next
            End If
            SelectSheet("sheet1")
            ClearArrays()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ClearArrays()
        Array.Clear(IL_XArray, 0, YArray.Count - 1)
        Array.Clear(IL1_YArray, 0, YArray.Count - 1)
        Array.Clear(IL2_YArray, 0, YArray.Count - 1)
        Array.Clear(RL_XArray, 0, YArray.Count - 1)
        Array.Clear(RL_YArray, 0, YArray.Count - 1)
        Array.Clear(DIR_XArray, 0, YArray.Count - 1)
        Array.Clear(DIR_YArray, 0, YArray.Count - 1)
        Array.Clear(COUP_XArray, 0, YArray.Count - 1)
        Array.Clear(COUP1_YArray, 0, YArray.Count - 1)
        Array.Clear(COUP2_YArray, 0, YArray.Count - 1)
        Array.Clear(ISO_XArray, 0, YArray.Count - 1)
        Array.Clear(ISO_YArray, 0, YArray.Count - 1)
        Array.Clear(AB_XArray, 0, YArray.Count - 1)
        Array.Clear(AB1_YArray, 0, YArray.Count - 1)
        Array.Clear(AB2_YArray, 0, YArray.Count - 1)
        Array.Clear(PB_XArray, 0, YArray.Count - 1)
        Array.Clear(PB1_YArray, 0, YArray.Count - 1)
        Array.Clear(PB2_YArray, 0, YArray.Count - 1)
    End Sub

    Private Sub txtOffset1_TextChanged(sender As Object, e As EventArgs) Handles txtOffset1.DoubleClick
        Dim SQLStr As String

        If Not IsNumeric(Me.txtOffset1.Text) Then Exit Sub
        SQLStr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
        If SQL.CheckforRow(SQLStr, "NetworkSpecs") = 0 Then
            SQLStr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
        End If

        SQLStr = "UPDATE Specifications Set Offset1 = " & Me.txtOffset1.Text & " where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
        SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")

    End Sub

    Private Sub txtOffset2_TextChanged(sender As Object, e As EventArgs) Handles txtOffset2.DoubleClick
        Dim SQLStr As String

        If Not IsNumeric(Me.txtOffset2.Text) Then Exit Sub
        SQLStr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
        If SQL.CheckforRow(SQLStr, "NetworkSpecs") = 0 Then
            SQLStr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
        End If

        SQLStr = "UPDATE Specifications Set Offset2 = " & Me.txtOffset2.Text & " where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
        SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
    End Sub

    Private Sub txtOffset3_TextChanged(sender As Object, e As EventArgs) Handles txtOffset3.DoubleClick
        Dim SQLStr As String

        If Not IsNumeric(Me.txtOffset3.Text) Then Exit Sub
        SQLStr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
        If SQL.CheckforRow(SQLStr, "NetworkSpecs") = 0 Then
            SQLStr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
        End If

        SQLStr = "UPDATE Specifications Set Offset3 = " & Me.txtOffset3.Text & " where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
        SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
    End Sub

    Private Sub txtOffset4_TextChanged(sender As Object, e As EventArgs) Handles txtOffset4.DoubleClick
        Dim SQLStr As String

        If Not IsNumeric(Me.txtOffset4.Text) Then Exit Sub
        SQLStr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
        If SQL.CheckforRow(SQLStr, "NetworkSpecs") = 0 Then
            SQLStr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
        End If

        SQLStr = "UPDATE Specifications Set Offset4 = " & Me.txtOffset4.Text & " where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
        SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
    End Sub

    Private Sub txtOffset5_TextChanged(sender As Object, e As EventArgs) Handles txtOffset5.DoubleClick
        Dim SQLStr As String

        If Not IsNumeric(Me.txtOffset5.Text) Then Exit Sub
        SQLStr = "SELECT * from Specifications where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
        If SQL.CheckforRow(SQLStr, "NetworkSpecs") = 0 Then
            SQLStr = "Insert Into Specifications (JobNumber, PartNumber) values ('" & Me.cmbJob.Text & "','" & Me.cmbPart.Text & "')"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
        End If

        SQLStr = "UPDATE Specifications Set Offset5 = " & Me.txtOffset5.Text & " where PartNumber = '" & Me.cmbPart.Text & "' And JobNumber = '" & Me.cmbJob.Text & "'"
        SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
    End Sub


    Private Sub DeleteOperatorToolStripMenuItem_Click(sender As Object, e As EventArgs)
        DeleteOp.Checked = True
        Dim OP As New OperatorEntry
        OP.StartPosition = FormStartPosition.Manual
        OP.Location = New Point(globals.XLocation, globals.YLocation)
        OP.ShowDialog()
    End Sub

    Private Sub ckTweakMode_CheckedChanged(sender As Object, e As EventArgs) Handles ckTweakMode.CheckedChanged
        TweakMode = ckTweakMode.Checked
        If Not TweakMode Then
            Data1.Text = ""
            Data2.Text = ""
            Data3.Text = ""
            If SpecAB_TF Then
                Data4L.Text = ""
                Data4H.Text = ""
            Else
                Data4.Text = ""
            End If
            Data5.Text = ""
            TEST1PASS = True
            TEST2PASS = True
            TEST3PASS = True
            If SpecAB_TF Then
                TEST4LPASS = True
                TEST4HPASS = True
            Else

                TEST4PASS = True
            End If
            TEST5PASS = True
            PF1.Text = "TBD"
            PF2.Text = "TBD"
            PF3.Text = "TBD"
            PF4.Text = "TBD"
            PF5.Text = "TBD"
            ClearFailureLog()
            CleanUpEffeciency()
            CheckTestResume()
            UUTNum = TempUUTNum
            StatusLog.Items.Add("Tweak Mode On:" & "" & DateTime.Now.ToString)
            Master_bypass = False
        Else
            ClearStatusLog()
            TempUUTNum = UUTNum
            StatusLog.Items.Add("Tweak Mode Off:" & "" & DateTime.Now.ToString)
            Master_bypass = True
        End If
    End Sub

    Private Sub ckROBOT_CheckedChanged(sender As Object, e As EventArgs) Handles ckROBOT.CheckedChanged
        If Me.ckROBOT.Checked And Not MCCLoaded Then LoadminiLAB()
        If Me.ckROBOT.Checked Then
            If FirstComplete Then
                TestCompleteSignal(False) ' Note False/False tells the Robot to start
                TestRunningSignal(True)
                Robot = True
                MCCLoaded = True
                ReadySignal = False
                FirstComplete = False
            End If
            saveConfigurationVal(iniPathName, "test_results", "TBD")
            txtFullAuto.Text = "Full Automation in Process. DO NOT TOUCH"
            If UUTCount.Text = 0 Then
                saveConfigurationVal(iniPathName, "uut_number", 1)
            Else
                saveConfigurationVal(iniPathName, "uut_number", UUTCount.Text)
            End If
            RobotTimer.Start()
        Else
            Robot = False
            txtFullAuto.Text = ""
            RobotTimer.Stop()
            TestCompleteSignal(False) ' Note False/False tells the Robot to stop
            TestRunningSignal(False)
            saveConfigurationVal(iniPathName, "test_results", "TBD")
        End If
    End Sub
    Private Sub RobotStatus()
        If Not TestRunning Then ReadySignal = GetReadySignal() ' From ROBOT
        RobotMoving = GetROBOTMovingSignal() ' From ROBOT
        RobotError = GetROBOTErrorSignal() ' From ROBOT
        Retest = GetRetest() ' From ROBOT
        GetTrayStatus()
        TestCompleteSignal(False)
        If NewTray.ToUpper.Contains("REPLACE") Then
            MsgBox("GIZMO says 'Replace the New Tray'", vbOK, "New Tray")
            saveConfigurationVal(iniPathName, "new_tray", "OK")
        End If
        If PassTray.ToUpper.Contains("REPLACE") Then
            MsgBox("GIZMO says 'Replace the Pass Tray'", vbOK, "Pass Tray")
            saveConfigurationVal(iniPathName, "pass_tray", "OK")
        End If
        If FailTray.ToUpper.Contains("REPLACE") Then
            MsgBox("GIZMO says 'Replace the Fail Tray'", vbOK, "Fail Tray")
            saveConfigurationVal(iniPathName, "fail_tray", "OK")
        End If

        If RobotError And ReadySignal And RobotMoving Then
            txtFullAuto.Text = "HALT!!! ROBOT is reporting an Error."
            MsgBox("GIZMO is reporting an error", vbOK, "ERROR.")

        ElseIf RobotError And (Not ReadySignal Or Not RobotMoving) Then
            txtFullAuto.Text = "ROBOT attention is required"
            MsgBox("GIZMO needs assistence", vbOK, "ERROR.")
        ElseIf ReadySignal And RobotMoving Then
            txtFullAuto.Text = "Full Automation. WARNING!!! ROBOT Moving"
        ElseIf ReadySignal And RobotMoving Then
            txtFullAuto.Text = "Full Automation. ROBOT Ready"
        Else
            txtFullAuto.Text = "Full Automation in Process. DO NOT TOUCH"
        End If
    End Sub

    Private Sub RobotTimer_Tick(sender As Object, e As EventArgs) Handles RobotTimer.Tick
        Dim Ready As Boolean = False
        If UUTCount.Text = Quantity Then
            RobotTimer.Stop()
            txtFullAuto.Text = "Full Automation Complete. Stopping ROBOT"
            TestCompleteSignal(False) ' Note False tells the Robot to stop
            TestRunningSignal(False) ' Note False tells the Robot to stop
            PassFailSignal("TBD")
        Else
            RobotStatus()
            If ReadySignal And Not RobotError And Me.cmdStartTest.Enabled And Not Me.cmdRetest.Enabled Then
                ReadySignal = False
                PassFailSignal("TBD")
                TestPassFail = False
                Me.cmdStartTest.PerformClick()
                TestRetest = 0
            ElseIf ReadySignal And Not TestRunning And Me.cmdStartTest.Enabled And Me.cmdRetest.Text = "Undo" And TestPassFail Then
                ReadySignal = False
                PassFailSignal("TBD")
                Me.cmdStartTest.PerformClick()
                TestRetest = 0
            ElseIf ReadySignal And Not RobotError And Not TestRunning And Me.cmdStartTest.Enabled And Me.cmdRetest.Enabled And Not TestPassFail And Retest Then
                ReadySignal = False
                PassFailSignal("TBD")
                Me.cmdRetest.PerformClick()
                TestRetest += 1
            ElseIf ReadySignal And Not RobotError And Not TestRunning And Me.cmdStartTest.Enabled And Me.cmdRetest.Enabled And Not TestPassFail And Not Retest Then
                ReadySignal = False
                PassFailSignal("TBD")
                Me.cmdStartTest.PerformClick()
                TestRetest = 0
            ElseIf TestComplete And Not TestRunning And Not RobotError And TestPassFail Then
                PassFailSignal(TestPassFail)
                TestCompleteSignal(True)
                TestRunningSignal(True)
                saveConfigurationVal(iniPathName, "uut_number", UUTCount.Text)
                TestComplete = False
            End If
        End If
    End Sub
    Private Sub ErrorTimer_Tick(sender As Object, e As EventArgs) Handles ErrorTimer.Tick
        txtFullAuto.Text = "TOO MANY FAILURES!!! Contact Supervisor"
        If Test1Failing Then
            status("Red", "TEST1")
            Me.Refresh()
            Delay(1000)
            txtFullAuto.Text = ""
            status("Black", "TEST1")
            Me.Refresh()
            Delay(1000)
        ElseIf Test2Failing Then
            status("Red", "TEST2")
            Me.Refresh()
            Delay(1000)
            txtFullAuto.Text = ""
            status("Black", "TEST2")
            Me.Refresh()
            Delay(1000)
        ElseIf TEST3Failing Then
            status("Red", "TEST3")
            Me.Refresh()
            Delay(1000)
            txtFullAuto.Text = ""
            status("Black", "TEST3")
            Me.Refresh()
            Delay(1000)
        ElseIf TEST4Failing Then
            status("Red", "TEST4")
            Me.Refresh()
            Delay(1000)
            txtFullAuto.Text = ""
            status("Black", "TEST4")
            Me.Refresh()
            Delay(1000)
        ElseIf TEST5Failing Then
            status("Red", "TEST5")
            Me.Refresh()
            Delay(1000)
            txtFullAuto.Text = ""
            status("Black", "TEST5")
            Me.Refresh()
            Delay(1000)
        Else
            status("Red", "TEST1")
            status("Red", "TEST2")
            status("Red", "TEST3")
            status("Red", "TEST4")
            status("Red", "TEST5")
            Me.Refresh()
            Delay(1000)
            txtFullAuto.Text = ""
            status("Black", "TEST1")
            status("Black", "TEST2")
            status("Black", "TEST3")
            status("Black", "TEST4")
            status("Black", "TEST5")
            Me.Refresh()
            Delay(1000)
        End If
    End Sub

    Private Sub MiniLabTestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MiniLabTestToolStripMenuItem.Click
        Dim OP As New CalGizmo
        OP.StartPosition = FormStartPosition.Manual
        OP.Location = New Point(globals.XLocation, globals.YLocation)
        OP.ShowDialog()
    End Sub

    Private Sub SupervisorPasswordToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SupervisorPasswordToolStripMenuItem.Click
        StatusLog.Items.Add("Opening Supervisor:" & "" & DateTime.Now.ToString)
        Dim SPEC As New ResetPassword
        SPEC.StartPosition = FormStartPosition.Manual
        ' SPEC.Location = New Point(Globals.XLocation + NewXLocation, Globals.YLocation + Globals.XSize)
        SPEC.ShowDialog()
    End Sub


    Private Sub txtArtwork_TextChanged(sender As Object, e As EventArgs) Handles txtArtwork.TextChanged
        UUTReset = True
        txtArtwork.SelectionStart = Len(txtArtwork.Text)
        txtArtwork.Text = Trim(txtArtwork.Text.ToUpper)
        ArtworkRevision = txtArtwork.Text

    End Sub


    Private Sub ReportServerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportServerToolStripMenuItem.Click
        Process.Start("microsoft-edge:http://inn-sqlexpress:8888/test/")
    End Sub

    Private Sub UpdateToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles UpdateToolStripMenuItem.Click
        Process.Start("\\ippdc\Test Automation\ATE_Publish\setup.exe")
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        Process.Start("http://inn-sqlexpress:8888/trouble/trouble_ticket_open/")
    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click

    End Sub
End Class
