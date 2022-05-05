Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection


Public Class frmMANUALTEST
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
    Private Reset As Boolean = False
    Private RunTest1 As Boolean = False
    Private RunTest2 As Boolean = False
    Private RunTest3 As Boolean = False
    Private RunTest4 As Boolean = False
    Private RunTest5 As Boolean = False
    Public SaveData As Boolean = False


    Public Sub New()
        InitializeComponent()
        Try
            SQLAccess = My.Computer.Network.Ping("INN-SQLEXPRESS")
        Catch
            SQLAccess = False
        End Try
        SQLVerified = SQLAccess
        Try
            TestRunningSignal(False)
            TestCompleteSignal(False)
            TestRunning = False
            NoInit = True
            NetworkAccess = CheckNetworkFolder()
            WorkStation = GetComputerName()
            If WorkStation = "mford" Then
                Simulation.Visible = True
            Else
                Simulation.Visible = False
                Pass.Visible = False
            End If
            Pass.Visible = False
            txtArtwork.Text = Artwork + Rev
            txtPanel.Text = Panel
            txtSector.Text = Sector
            txtLOT.Text = LOT
            If txtLOT.Text.Length() = 1 Then
                txtLOT.Text = "000000000000" + txtLOT.Text
            ElseIf txtLOT.Text.Length() = 2 Then
                txtLOT.Text = "00000000000" + txtLOT.Text
            ElseIf txtLOT.Text.Length() = 3 Then
                txtLOT.Text = "0000000000" + txtLOT.Text
            ElseIf txtLOT.Text.Length() = 4 Then
                txtLOT.Text = "000000000" + txtLOT.Text
            ElseIf txtLOT.Text.Length() = 5 Then
                txtLOT.Text = "00000000" + txtLOT.Text
            ElseIf txtLOT.Text.Length() = 6 Then
                txtLOT.Text = "0000000" + txtLOT.Text
            ElseIf txtLOT.Text.Length() = 7 Then
                txtLOT.Text = "000000" + txtLOT.Text
            ElseIf txtLOT.Text.Length() = 8 Then
                txtLOT.Text = "00000" + txtLOT.Text
            ElseIf txtLOT.Text.Length() = 9 Then
                txtLOT.Text = "0000" + txtLOT.Text
            ElseIf txtLOT.Text.Length() = 10 Then
                txtLOT.Text = "000" + txtLOT.Text
            ElseIf txtLOT.Text.Length() = 11 Then
                txtLOT.Text = "00" + txtLOT.Text
            ElseIf txtLOT.Text.Length() = 12 Then
                txtLOT.Text = "0" + txtLOT.Text
            End If
            LOT = txtLOT.Text
            Version = "Version: " & GetVersion()
            xTitle = "Freq MHz"
            yTitle = "Signal dBm"
            Temperature = "25"
            CalDue = Now
            SupervisorPassword = SQL.GetPassword()
            Dim SQLstr As String

            Me.txtVersion.Text = "Version: " + GetVersion()
            Uploading = False
            ILSetDone = False

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
            SwitchModel = GetSwitchModel()
            If SwitchModel = "RC-1SP6T-A12" Then
                btSwitch1.Visible = True
                btSwitch2.Visible = True
                btSwitch3.Visible = True
                btSwitch4.Visible = True
                btSwitch5.Visible = True
                btSwitch6.Visible = True
                cmbSwitch.Text = "Switch POS 1"
            Else
                btSwitch1.Visible = True
                btSwitch2.Visible = True
                btSwitch3.Visible = True
                btSwitch4.Visible = True
                cmbSwitch.Text = "Switch POS 1"
            End If
        Catch ex As Exception

        End Try
    End Function
    Private Sub cmbJob_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbJob.SelectedIndexChanged
        Try
            If Me.cmbJob.GetItemText(Me.cmbJob.SelectedItem) <> " " And Uploading = False Then
                If Not DontclickTheButton Then
                    If Not Connected Then
                        ScanGPIB.connect("GPIB0::16::INSTR", GetTimeout())
                        Connected = True
                    End If
                    Reset = True
                    Dim SwNum As Integer
                    Dim SerList As String = ""
                    Dim Firm As Integer
                    ManualTests.InitializeSwitch(SwNum, SerList, Firm)
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

                    Me.Refresh()
                    'PartClicked()
                    DontclickTheButton = False

                    SwitchPorts = SQL.GetSpecification("SwitchPorts")
                    LastJob = Job
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub JobClicked()
        Dim SQLstr As String
        Dim SwPos As String
        Dim GoodJob As Boolean
        Try

            If Me.cmbJob.Text = "Add New" Then
                Dim SPEC As New frmSpecifications
                SPEC.StartPosition = FormStartPosition.Manual
                ' SPEC.Location = New Point(Globals.XLocation + NewXLocation, Globals.YLocation + Globals.XSize)
                SPEC.ShowDialog()
                Me.Hide()
                frmSpecifications.ShowDialog()
            End If
            txtArtwork.Text = Artwork + Rev
            txtPanel.Text = Panel
            txtSector.Text = Sector
            txtLOT.Text = LOT
            Panel = txtPanel.Text
            Sector = txtSector.Text
            ArtworkRevision = Artwork + Rev + Panel + Sector + LOT
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
                        txtKitQty.Text = dr.Item(5)
                        If dr.Item(1) = "90 DEGREE COUPLER" Or dr.Item(1) = "90 DEGREE COUPLER SMD" Then
                            If dr.Item(1) = "90 DEGREE COUPLER SMD" Then
                                SMD = True
                            Else
                                SMD = False
                            End If
                            SpecIndex = 0
                            SpecType = "90 DEGREE COUPLER"
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True

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

                            TestLabel3.Visible = False
                            Spec3Min.Visible = False
                            Spec3Max.Visible = False
                            Data3.Visible = False
                            PF3.Visible = False
                            txtOffset3.Visible = False
                            TestLabel4.Visible = False
                            Spec4Min.Visible = False
                            Spec4Max.Visible = False
                            Data4.Visible = False
                            PF4.Visible = False
                            txtOffset4.Visible = False
                            TestLabel5.Visible = False
                            Spec5Min.Visible = False
                            Spec5Max.Visible = False
                            Data5.Visible = False
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
                            TestLabel3.Visible = False
                            Spec3Min.Visible = False
                            Spec3Max.Visible = False
                            Data3.Visible = False
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
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
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
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
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
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
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
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
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
                        TestLabel3.Visible = True
                        Spec3Min.Visible = True
                        Spec3Max.Visible = True
                        Data3.Visible = True
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
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
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
                            TestLabel3.Visible = False
                            Spec3Min.Visible = False
                            Spec3Max.Visible = False
                            Data3.Visible = False
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
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
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
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
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
                            TestLabel3.Visible = True
                            Spec3Min.Visible = True
                            Spec3Max.Visible = True
                            Data3.Visible = True
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
                    MYMsgBox(SpecStopFreq & "MHz  Exceeds the Frequency Range of the " & Me.cmbVNA.Text & ".  Please move to capible workstation or choose another Job")
                    Exit Sub
                End If
            End If

            If Me.cmbVNA.Text = "HP_8753C" Then
                If SpecStopFreq > GetVNAFreq() Then
                    MYMsgBox(SpecStopFreq & "MHz  Exceeds the Calibrated Frequency Range of the " & Me.cmbVNA.Text & ".  Please Calibrate the 6GHZ Range")
                    ManualTests.CalibrateVNA()
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

            UUTMessage.Text = "  UUT ManualTests  --  Load Unit #1"


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

                UUTMessage.Text = "  UUT ManualTests  --  Load Part Number"

            End If

            ResetManualTests(, resumeTest)
            If Not resumeTest Then
                UUTNum = 0
                UUTStatusColor.BackColor = Color.LawnGreen
            End If

            If Simulation.Checked = False And GoodJob Then
                RecallCal(1)
                SetupVNA(True, 1)
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
                            TestLabel3.Enabled = False
                            Spec3Min.Enabled = False
                            Spec3Max.Enabled = False
                            Data3.Enabled = False
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
                            TestLabel3.Enabled = False
                            Spec3Min.Enabled = False
                            Spec3Max.Enabled = False
                            Data3.Enabled = False
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
                    MYMsgBox(SpecStopFreq & "MHz  Exceeds the Frequency Range of the " & Me.cmbVNA.Text & ".  Please move to capible workstation or choose another Job")
                    Exit Sub
                End If
            End If

            If Me.cmbVNA.Text = "HP_8753C" Then
                If SpecStopFreq > GetVNAFreq() Then
                    MYMsgBox(SpecStopFreq & "MHz  Exceeds the Calibrated Frequency Range of the " & Me.cmbVNA.Text & ".  Please Calibrate the 6GHZ Range")
                    ManualTests.CalibrateVNA()
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



            UUTMessage.Text = "  UUT ManualTests  --  Load Unit #1"



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
                UUTMessage.Text = "  UUT ManualTests  --  Load Part Number"

            End If

            ResetManualTests()
            If Not SpecType = Nothing Then
                'MSChart.InitializeChart()
                'MSChart.ResetChartData(SpecType)
                'MSChart.UpDateChart(SpecType)
                UUTNum = 0
                UUTStatusColor.BackColor = Color.LawnGreen
            End If

            If Simulation.Checked = False And GoodJob Then SetupVNA(True, 1)
            Exit Sub

        Catch ex As Exception

        End Try


    End Sub

    Private Sub GetTrace_CheckedChanged(sender As Object, e As EventArgs) Handles GetTrace.CheckedChanged
        If Me.GetTrace.CheckState = CheckState.Checked Then
            GetTrace.Text = "Trace test mode"
        Else
            GetTrace.Text = "Marker test mode"
        End If
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

    Private Sub ResetManualTests(Optional Retest As Boolean = False, Optional resumeTest As Boolean = False)
        Try
            If resumeTest Then Exit Sub
            If Not Retest Then UUTFail = 0

            If Not Retest Then
                TEST1FailRetest = 0
                PF1.ForeColor = Color.CornflowerBlue
                PF1.Text = "TBD"
                Data1.Text = ""
            ElseIf Retest Then
                PF1.ForeColor = Color.CornflowerBlue
                PF1.Text = "TBD"
                Data1.Text = ""
            End If

            If Not Retest Then
                TEST2FailRetest = 0
                PF2.ForeColor = Color.CornflowerBlue
                PF2.Text = "TBD"
                Data2.Text = ""
            ElseIf Retest Then
                PF2.ForeColor = Color.CornflowerBlue
                PF2.Text = "TBD"
                Data2.Text = ""
            End If

            If Not Retest Then
                TEST3FailRetest = 0
                PF3.ForeColor = Color.CornflowerBlue
                PF3.Text = "TBD"
                Data3.Text = ""
            ElseIf Retest Then
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
            ElseIf Retest Then
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
            ElseIf Retest Then
                PF5.ForeColor = Color.CornflowerBlue
                PF5.Text = "TBD"
                Data5.Text = ""
            End If

            If Not resumeTest Then Me.UUTStatusColor.BackColor = Color.LawnGreen

        Catch
        End Try

    End Sub

    Private Sub Simulation_Changed(sender As Object, e As EventArgs) Handles Simulation.CheckedChanged
        If Simulation.Checked Then
            Me.RealData.Visible = True
            Me.Pass.Visible = True
            Me.Fail.Visible = True
            DebugChecked = True
        Else
            Me.RealData.Visible = False
            Me.Pass.Visible = False
            Me.Fail.Visible = False
            DebugChecked = False
        End If
        Debug = Simulation.Checked

    End Sub
    Private Sub AllManualTests(TestRun As String)
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
            resumeTest = False

            StartTime = Now()
            TestRunning = True
            TestComplete = False
            GlobalFailed = False

            txtArtwork.Text = Artwork + Rev
            txtPanel.Text = Panel
            txtSector.Text = Sector
            txtLOT.Text = LOT
            Panel = txtPanel.Text
            Sector = txtSector.Text
            ArtworkRevision = Artwork + Rev + Panel + Sector + LOT

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

            ResetManualTests(, resumeTest)
            GetLoss()
            SwPos = ""
            If DontclickTheButton = True Then Exit Sub
            If Me.cmbJob.Text = "" Or Me.cmbJob.Text = " " Then
                MYMsgBox("Please select Job")
                Exit Sub
            End If
            If Me.txtArtwork.Text = "" Then
                MYMsgBox("Please input the Artwork Revision")
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
                MYMsgBox("Please Choose the Job Number first", , "Not Ready to Start")
                Exit Sub
            Else
                TestExist = False
                If resumeTest Then
                    LastTest = Date.Now.Ticks
                    PPH = SQL.GetSpecification("PartsPerHour")
                    Quantity = SQL.GetSpecification("Quantity")
                    Dim OP As New OperatorEntry
                    OP.StartPosition = FormStartPosition.Manual
                    OP.Location = New Point(globals.XLocation + 450, globals.YLocation + 200)
                    OP.ShowDialog()

                    'ATS Developer does not save data while developing
                    If User = "ATS" Then
                        GetTrace.CheckState = CheckState.Unchecked
                    End If
                End If
            End If
            If Reset Then
                ResetManualTests()
                Reset = False
            End If

            UUTFail = 0
            Data1.Text = ""
            Data1L.Text = ""
            Data1H.Text = ""
            Data2.Text = ""
            Data3.Text = ""
            Data4.Text = ""
            Data4L.Text = ""
            Data4H.Text = ""
            Data5.Text = ""
            If UUTNum = 0 Or Resumed Then UUTNum = UUTNum + 1

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

            If TweakMode Then UUTNum = 1
            If UUTNum_Reset > 5 And Not BypassUnchecked Then
                Me.GetTrace.Checked = False
            ElseIf UUTNum_Reset > 5 And BypassUnchecked Then
                Me.GetTrace.Checked = True
            End If

            If TweakMode Then
                Me.UUTMessage.Text = "  UUT ManualTests  --   Testing Undisclosed Unit "
            ElseIf Not TraceChecked Then
                UUTMessage.Text = "  UUT ManualTests Marker Mode  --   Load Unit #" & UUTNum + 1
            Else
                Me.UUTMessage.Text = "  UUT ManualTests  --   Testing Unit #" & UUTNum
            End If

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
            If RunTest1 Then
                If TraceChecked Then
                    If (SpecType = "TRANSFORMER") And IL_TF Then PassFail = ManualTests.InsertionLossTRANS_multiband(TestRun, , TestID)
                    If SpecType = "TRANSFORMER" And Not IL_TF Then PassFail = ManualTests.InsertionLossTRANS(TestRun, , TestID)
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType = "BALUN") And SpecAB_TF Then PassFail = ManualTests.InsertionLoss3dB_multiband(TestRun, , TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = ManualTests.InsertionLoss3dB(TestRun, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = ManualTests.InsertionLossCOMB(TestRun, , TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = ManualTests.InsertionLossDIR(TestRun, , TestID)
                ElseIf Not MutiCalChecked Then
                    If (SpecType = "TRANSFORMER") And IL_TF Then PassFail = ManualTests.InsertionLossTRANS_multiband(TestRun, , TestID)
                    If SpecType = "TRANSFORMER" And Not IL_TF Then PassFail = ManualTests.InsertionLossTRANS_Marker(TestRun, , TestID)
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And SpecAB_TF Then PassFail = ManualTests.InsertionLoss3dB_multiband(TestRun, , TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = ManualTests.InsertionLoss3dB_marker(TestRun, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = ManualTests.InsertionLossCOMB_Marker(TestRun, , TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = ManualTests.InsertionLossDIR_Marker(TestRun, , TestID)
                End If


                If IL_TF Then
                    RetrnVal1 = IL1
                    If SaveData Then SaveTestData("Manual " & TestRun & "Manual " & TestRun & "InsertionLoss", RetrnVal1, 1)
                    RetrnVal = IL2
                    If SaveData Then SaveTestData("Manual " & TestRun & "Manual " & TestRun & "InsertionLoss2", RetrnVal, 1)
                    Data1L.Text = IL1
                    Data1H.Text = IL2
                Else
                    RetrnVal = IL
                    If SaveData Then SaveTestData("Manual " & TestRun & "InsertionLoss", RetrnVal, 1)
                    RetrnStr = CStr(TruncateDecimal(RetrnVal, 2))
                    Data1.Text = Format(RetrnVal, "0.00")
                End If
                status("Blue", "TEST1")
                PF1.Text = PassFail


                If PassFail = "Pass" Then
                    TEST1PASS = True
                    status("Green", "TEST1")
                ElseIf PassFail = "Fail" Then
                    TEST1PASS = False
                    status("Red", "TEST1")
                    Test1Fail = Test1Fail + 1
                    If Not GlobalFailed Then
                        GlobalFail = GlobalFail + 1
                        GlobalFailed = True
                    End If

                    status("Red", "TEST1")
                    TEST1FailRetest = TEST1FailRetest + 1
                    UUTFail = 1
                End If
            Else
                TEST1PASS = True
                status("Blue", "TEST1")
                SaveTestData("Manual " & TestRun & "InsertionLoss", GetSpecification("InsertionLoss"), 1)
            End If
            Me.Refresh()

            'Return Loss
            If RunTest2 Then
                If TraceChecked And Not TweakMode Then
                    PassFail = ManualTests.ReturnLoss(TestRun, , TestID)
                Else
                    PassFail = ManualTests.ReturnLoss_Marker(TestRun, , TestID)
                End If
                RetrnVal = RL
                SaveTestData("Manual " & TestRun & "ReturnLoss", RetrnVal, 2)
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
                    status("Red", "TEST2")
                    TEST2FailRetest = TEST2FailRetest + 1
                    UUTFail = 1
                End If
                If SaveData Then SaveTestData("Manual " & TestRun & "ReturnLoss_Manual", VSWRtoRL(SQL.GetSpecification("VSWR")), 2)
            Else
                TEST2PASS = True
                status("Blue", "TEST2")
            End If
            Me.Refresh()
            If SpecType <> "COMBINER/DIVIDER" Then GoTo Test2Sub
Test2SubRet:

            'AmplitudeBalance
            'Directivity
            If RunTest4 Then
                If TraceChecked Then
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And SpecAB_TF Then PassFail = ManualTests.AmplitudeBalance_multiband(TestRun, , TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = ManualTests.AmplitudeBalance(TestRun, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = ManualTests.AmplitudeBalanceCOMB(TestRun, , TestID)
                    If SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Then PassFail = ManualTests.Directivity(TestRun, 1, SpecType, , TestID)
                ElseIf Not MutiCalChecked Then
                    If (SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN")) And SpecAB_TF Then PassFail = ManualTests.AmplitudeBalance_multiband(TestRun, , TestID)
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = ManualTests.AmplitudeBalance_Marker(TestRun, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = ManualTests.AmplitudeBalanceCOMB_Marker(TestRun, , TestID)
                    If SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Then PassFail = ManualTests.Directivity_Marker(TestRun, 1, SpecType, , TestID)
                End If
                If SpecType <> "SINGLE DIRECTIONAL COUPLER" Then
                    If SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Then
                        RetrnVal = DIR
                        If SaveData Then SaveTestData("Manual " & TestRun & "Directivity", RetrnVal, 3)
                    Else
                        If SpecAB_TF Then
                            RetrnVal = AB1
                            If SaveData Then SaveTestData("Manual " & TestRun & "AmplitudeBalance1", RetrnVal, 3)
                            RetrnVal = AB2
                            If SaveData Then SaveTestData("Manual " & TestRun & "AmplitudeBalance2", RetrnVal, 3)
                            'remove later
                            RetrnVal = AB
                            If SaveData Then SaveTestData("Manual " & TestRun & "AmplitudeBalance", RetrnVal, 3)
                        Else
                            RetrnVal = AB
                            If SaveData Then SaveTestData("Manual " & TestRun & "AmplitudeBalance", RetrnVal, 3)
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
                            TEST4Fail = TEST4Fail + 1
                            GlobalFail = GlobalFail + 1
                            TEST4PASS = False
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
                        End If
                    End If

                End If
            End If
            Me.Refresh()
            'PhaseBalance
            'CoupledFlatness
            If RunTest5 Then
                If TraceChecked Then
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = ManualTests.PhaseBalance(TestRun, SpecType, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = ManualTests.PhaseBalanceCOMB(TestRun, SpecType, , TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = ManualTests.CoupledFlatness(TestRun, 1, SpecType, , TestID)
                ElseIf Not MutiCalChecked Then
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = ManualTests.PhaseBalance_Marker(TestRun, SpecType, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = ManualTests.PhaseBalanceCOMB_Marker(TestRun, SpecType, , TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = ManualTests.CoupledFlatness(TestRun, 1, SpecType, , TestID)
                End If
                If InStr(SpecType, "DIRECTIONAL COUPLER") Then
                    RetrnVal = CF
                    If SaveData Then SaveTestData("Manual " & TestRun & "CoupledFlatness", RetrnVal, 5)
                Else
                    RetrnVal = PB
                    If SaveData Then SaveTestData("Manual " & TestRun & "PhaseBalance", RetrnVal, 5)
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
                        status("Red", "TEST5")
                        TEST5FailRetest = TEST5FailRetest + 1
                        UUTFail = 1
                    End If
                End If
            Else
                TEST5PASS = True
                status("Blue", "TEST5")
                If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then If SaveData Then SaveTestData("Manual " & TestRun & "PhaseBalance", GetSpecification("PhaseBalance"), 5)
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then If SaveData Then SaveTestData("Manual " & TestRun & "CoupledFlatness", GetSpecification("CoupledFlatness"), 5)
            End If

            If SpecType <> "COMBINER/DIVIDER" Then
                If SpecType = "DUAL DIRECTIONAL COUPLER" And PF1.Text <> "Fail" And PF2.Text <> "Fail" And (PF3.Text = "Fail" Or PF4.Text = "Fail" Or PF5.Text = "Fail") Then
                    If SaveData Then
                        If MYMsgBox("Try Forward Measurement Again", vbYesNo) = vbYes Then
                            GoTo Test2Sub
                        Else
                            GoTo TestComplete
                        End If
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
            If RunTest3 Then
                If TraceChecked Then
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = ManualTests.Isolation(TestRun, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = ManualTests.IsolationCOMB(TestRun, , TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = ManualTests.Coupling(TestRun, 1, SpecType, , TestID)
                ElseIf Not MutiCalChecked Then
                    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Then PassFail = ManualTests.Isolation_Marker(TestRun, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") And ISO_TF Then PassFail = ManualTests.IsolationCOMB(TestRun, , TestID)
                    If SpecType.Contains("COMBINER/DIVIDER") Then PassFail = ManualTests.IsolationCOMB_Marker(TestRun, , TestID)
                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then PassFail = ManualTests.Coupling(TestRun, 1, SpecType, , TestID)
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
                            If SaveData Then SaveTestData("Manual " & TestRun & "IsolationL", ISoL, 3)
                            RetrnStr = CStr(TruncateDecimal(ISoL, 1))
                            If SaveData Then SaveTestData("Manual " & TestRun & "IsolationH", ISoH, 3)
                            RetrnStr = CStr(TruncateDecimal(ISoH, 1))
                        Else
                            If SaveData Then SaveTestData("Manual " & TestRun & "IsoL", ISoL, 3)
                            RetrnStr = CStr(TruncateDecimal(ISoL, 1))
                            If SaveData Then SaveTestData("Manual " & TestRun & "IsoH", ISoH, 3)
                            RetrnStr = CStr(TruncateDecimal(ISoH, 1))
                        End If
                    ElseIf SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("COMBINER/DIVIDER") Or SpecType.Contains("BALUN") Then
                        RetrnVal = ISo
                        If SQLAccess Then
                            If SaveData Then SaveTestData("Manual " & TestRun & "Isolation", RetrnVal, 3)
                            RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                        Else
                            If SaveData Then SaveTestData("Manual " & TestRun & "Iso", RetrnVal, 3)
                            RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                        End If
                    Else
                        RetrnVal = COuP
                        If SaveData Then SaveTestData("Manual " & TestRun & "Coupling", RetrnVal, 3)
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
                                status("Red", "TEST3")
                                TEST3FailRetest = TEST3FailRetest + 1
                                UUTFail = 1
                            End If
                        End If
                    End If
                Else
                    TEST3PASS = True
                    status("Blue", "TEST3")
                    If SQLAccess Then
                        If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") And SQLAccess Then If SaveData Then SaveTestData("Manual " & TestRun & "Isolation", 0 - GetSpecification("Isolation"), 3)
                        If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") And Not SQLAccess Then If SaveData Then SaveTestData("Manual " & TestRun & "Isolation", 0 - GetSpecification("Isolation"), 3)
                    Else
                        If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") And SQLAccess Then If SaveData Then SaveTestData("Manual " & TestRun & "Iso", 0 - GetSpecification("Isolation"), 3)
                        If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") And Not SQLAccess Then If SaveData Then SaveTestData("Manual " & TestRun & "Iso", 0 - GetSpecification("Isolation"), 3)
                    End If

                    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then If SaveData Then SaveTestData("Manual " & TestRun & "Coupling", 0 - GetSpecification("Coupling"), 3)
                End If
                Me.Refresh()
            End If

            If Not SpecType.Contains("COMBINER/DIVIDER") Then GoTo Test2SubRet

TestComplete:  ' For everything except Directional Couplers

            'Directonal Couplers reverse direction
            If Not TweakMode And (SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or (SpecType = "SINGLE DIRECTIONAL COUPLER" And RunTest4)) Then
                If SpecType = "DUAL DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, REFL = SW3"
                If SpecType = "DUAL DIRECTIONAL COUPLER" And SwitchPorts = 1 Then SwPos = " OUT = SW1, CPL_J4 = SW2, ISO_J3 = SW3, CPL_J3 = SW4, ISO_J4 = SW5"
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, ISO = SW3"
                If SpecType = "BI DIRECTIONAL COUPLER" Then SwPos = "          OUT = SW1, CPL = SW2, REFL = SW3"
                txtTitle.Text = SpecType & SwPos
                If SaveData Then
                    MYMsgBox("Please turn the Directional Coupler in the Reverse direction")
                Else
                    GoTo TestReallyComplete
                End If

            End If

            'Reverse Coupling
            If SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" And RunTest3 Then
                PassFail = ManualTests.Coupling(TestRun, 2, SpecType, , TestID)
                RetrnVal = COuP
                RetrnStr = CStr(TruncateDecimal(RetrnVal, 1))
                If SaveData Then SaveTestData("Manual " & TestRun & "Coupling", RetrnVal, 3)
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
                End If
            End If
            Me.Refresh()

            ' Reverse Directivity
            If (SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "SINGLE DIRECTIONAL COUPLER") And RunTest4 Then
                If TraceChecked And Not TweakMode Then
                    PassFail = ManualTests.Directivity(TestRun, 2, SpecType, , TestID)
                Else
                    PassFail = ManualTests.Directivity_Marker(TestRun, 2, SpecType, , TestID)
                End If

                status("Blue", "TEST4")
                PF4.Text = PassFail
                RetrnVal = DIR
                If SaveData Then SaveTestData("Manual " & TestRun & "Directivity", RetrnVal, 4)
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
                End If
            End If
            Me.Refresh()
            ' Reverse Coupled Flatness
            If SpecType = "DUAL DIRECTIONAL COUPLER" And RunTest5 Then
                PassFail = ManualTests.CoupledFlatness(TestRun, 2, SpecType, , TestID)
                RetrnVal = CF
                SaveTestData("Manual " & TestRun & "CoupledFlatness", RetrnVal, 5)
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
                    status("Red", "TEST5")
                    TEST5FailRetest = TEST5FailRetest + 1
                    MSChart.UpDateChartData(SpecType, "CB", "Fail")
                    UUTFail = 1
                End If
            End If
            If SpecType = "DUAL DIRECTIONAL COUPLER" And Not Dir1Failed And PF1.Text <> "Fail" And PF2.Text <> "Fail" And (PF3.Text = "Fail" Or PF4.Text = "Fail" Or PF5.Text = "Fail") Then
                If SaveData Then
                    If MYMsgBox("Try Reverse Measurement Again?", vbYesNo) = vbYes Then
                        GoTo TestComplete
                    End If
                End If

            End If


            If Not SaveData Then
                UUTMessage.Text = "  UUT ManualTests  --   Tweak Mode. No Data Logging"
            ElseIf Not TraceChecked Then
                UUTMessage.Text = "  UUT ManualTests Marker Mode  --   Load Unit #" & UUTNum + 1
            Else
                UUTMessage.Text = "  UUT ManualTests  --   Load Unit #" & UUTNum + 1
            End If
TestReallyComplete:
            'Change Title Back To Forward
            If SpecType = "DUAL DIRECTIONAL COUPLER" Or (SpecType = "SINGLE DIRECTIONAL COUPLER" And RunTest4) Then
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

            ' Me.MasterUpLoad.Enabled = True
        Catch ex As Exception

        End Try
    End Sub
    Private Sub ResetLot()
        'MSChart.ResetChartData(SpecType)
        'MSChart.UpDateChart(SpecType)
        Me.UUTStatusColor.BackColor = Color.LawnGreen
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
        Me.PF2.ForeColor = Color.CornflowerBlue
        Me.PF2.Text = "TBD"
        Me.Data2.Text = ""
        If Not SpecType = "TRANSFORMER" Then
            Me.Data3.Visible = True
            Me.Spec3Min.Visible = True
            Me.Spec3Max.Visible = True
            Me.TestLabel3.Visible = True
            Me.txtOffset3.Visible = True
            Me.PF3.Visible = True
            Me.Data4.Visible = True
            Me.Spec4Min.Visible = True
            Me.Spec4Max.Visible = True
            Me.TestLabel4.Visible = True
            Me.txtOffset4.Visible = True
            Me.PF4.Visible = True
            Me.Data5.Visible = True
            Me.Spec5Min.Visible = True
            Me.Spec5Max.Visible = True
            Me.TestLabel5.Visible = True
            Me.txtOffset5.Visible = True
            Me.PF5.Visible = True
        End If
        Me.PF3.ForeColor = Color.CornflowerBlue
        Me.PF3.Text = "TBD"
        Me.Data3.Text = ""

        Me.PF4.ForeColor = Color.CornflowerBlue
        Me.PF4.Text = "TBD"
        Me.Data4.Text = ""

        Me.PF5.ForeColor = Color.CornflowerBlue
        Me.PF5.Text = "TBD"
        Me.Data5.Text = ""
    End Sub

    Public Sub Switch(index As Integer)
        Dim status As Integer
        Dim StatusRet As Integer

        SwitchCom.Connect()
        ILSetDone = False

        If index = 0 Then
            cmbSwitch.Text = "Switch POS 1"
            status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 1:" & "" & DateTime.Now.ToString)
        ElseIf index = 1 Then
            cmbSwitch.Text = "Switch POS 3"
            status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 2:" & "" & DateTime.Now.ToString)
        ElseIf index = 2 Then
            cmbSwitch.Text = "Switch POS 3"
            status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 3:" & "" & DateTime.Now.ToString)
        ElseIf index = 3 Then
            cmbSwitch.Text = "Switch POS 4"
            status = SwitchCom.SetSwitchPosition(4) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 4:" & "" & DateTime.Now.ToString)
        ElseIf index = 4 Then
            cmbSwitch.Text = "Switch POS 5"
            status = SwitchCom.SetSwitchPosition(5) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 4:" & "" & DateTime.Now.ToString)
        ElseIf index = 5 Then
            cmbSwitch.Text = "Switch POS 6"
            status = SwitchCom.SetSwitchPosition(6) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            System.Threading.Thread.Sleep(500)
            ' StatusLog.Items.Add("Switch POS 4:" & "" & DateTime.Now.ToString)
        End If
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
    Private Sub status(Color As String, Test As String, Optional Retest As Boolean = False)
        If Test = "TEST1" Then
            If Color = "Green" Then
                If UUTFail = 0 And Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                Me.PF1.ForeColor = Drawing.Color.LawnGreen
                If IL_TF Then
                    Me.Data1L.ForeColor = Drawing.Color.LawnGreen
                    Me.Data1H.ForeColor = Drawing.Color.LawnGreen
                Else
                    Me.Data1.ForeColor = Drawing.Color.LawnGreen
                End If
            End If

            If Color = "Red" Then
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
                Me.PF1.ForeColor = Drawing.Color.Black
            End If

        End If

        If Test = "TEST2" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                Me.PF2.ForeColor = Drawing.Color.LawnGreen
                Me.Data2.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
                Me.PF2.ForeColor = Drawing.Color.Red
                Me.Data2.ForeColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF2.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF2.ForeColor = Drawing.Color.Black
                Me.Data2.ForeColor = Drawing.Color.Black
                Me.PF2.ForeColor = Drawing.Color.Black
            End If
        End If
        If Test = "TEST3" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                Me.PF3.ForeColor = Drawing.Color.LawnGreen
                Me.Data3.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                Me.PF3.ForeColor = Drawing.Color.Red
                Me.Data3.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF3.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF3.ForeColor = Drawing.Color.Black
                Me.Data3.ForeColor = Drawing.Color.Black
                Me.PF3.ForeColor = Drawing.Color.Black
            End If
        End If
        If Test = "TEST3L" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                Me.PF3.ForeColor = Drawing.Color.LawnGreen
                Me.Data3L.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                Me.PF3.ForeColor = Drawing.Color.Red
                Me.Data3L.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF3.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF3.ForeColor = Drawing.Color.Black
                Me.Data3L.ForeColor = Drawing.Color.Black
                Me.PF3.ForeColor = Drawing.Color.Black
            End If
        End If
        If Test = "TEST3H" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                Me.PF3.ForeColor = Drawing.Color.LawnGreen
                Me.Data3H.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                Me.PF3.ForeColor = Drawing.Color.Red
                Me.Data3H.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF3.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF3.ForeColor = Drawing.Color.Black
                Me.Data4H.ForeColor = Drawing.Color.Black
                Me.PF3.ForeColor = Drawing.Color.Black
            End If
        End If

        If Test = "TEST4" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                Me.PF4.ForeColor = Drawing.Color.LawnGreen
                Me.Data4.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                Me.PF4.ForeColor = Drawing.Color.Red
                Me.Data4.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF4.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF4.ForeColor = Drawing.Color.Black
                Me.Data4.ForeColor = Drawing.Color.Black
                Me.PF4.ForeColor = Drawing.Color.Black
            End If
        End If
        If Test = "TEST4L" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                Me.PF4.ForeColor = Drawing.Color.LawnGreen
                Me.Data4L.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                Me.PF4.ForeColor = Drawing.Color.Red
                Me.Data4L.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF4.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF4.ForeColor = Drawing.Color.Black
                Me.Data4L.ForeColor = Drawing.Color.Black
            End If
        End If
        If Test = "TEST4H" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                Me.PF4.ForeColor = Drawing.Color.LawnGreen
                Me.Data4H.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                Me.PF4.ForeColor = Drawing.Color.Red
                Me.Data4H.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF4.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF4.ForeColor = Drawing.Color.Black
                Me.Data4H.ForeColor = Drawing.Color.Black
                Me.PF4.ForeColor = Drawing.Color.Black
            End If
        End If

        If Test = "TEST5" Then
            If Color = "Green" Then
                If UUTFail = 0 Then Me.UUTStatusColor.BackColor = Drawing.Color.LawnGreen
                Me.PF5.ForeColor = Drawing.Color.LawnGreen
                Me.Data5.ForeColor = Drawing.Color.LawnGreen
            End If
            If Color = "Red" Then
                Me.PF5.ForeColor = Drawing.Color.Red
                Me.Data5.ForeColor = Drawing.Color.Red
                Me.UUTStatusColor.BackColor = Drawing.Color.Red
            End If
            If Color = "Blue" Then
                PF5.ForeColor = Drawing.Color.CornflowerBlue
            End If
            If Color = "Black" Then
                PF5.ForeColor = Drawing.Color.Black
                Me.Data5.ForeColor = Drawing.Color.Black
                Me.PF5.ForeColor = Drawing.Color.Black
            End If
        End If
    End Sub



    Private Sub cmdStartTest1_Click(sender As Object, e As EventArgs) Handles cmdStartTest1.Click
        SaveData = False
        RunTest1 = True
        RunTest2 = False
        RunTest3 = False
        RunTest4 = False
        RunTest5 = False
        AllManualTests("Tuning")
    End Sub
    Private Sub cmdStartTest2_Click(sender As Object, e As EventArgs) Handles cmdStartTest2.Click
        SaveData = False
        RunTest1 = False
        RunTest2 = True
        RunTest3 = False
        RunTest4 = False
        RunTest5 = False
        AllManualTests("Tuning")
    End Sub
    Private Sub cmdStartTest3_Click(sender As Object, e As EventArgs) Handles cmdStartTest3.Click
        SaveData = False
        RunTest1 = False
        RunTest2 = False
        RunTest3 = True
        RunTest4 = False
        RunTest5 = False
        AllManualTests("Tuning")
    End Sub
    Private Sub cmdStartTest4_Click(sender As Object, e As EventArgs) Handles cmdStartTest4.Click
        SaveData = False
        RunTest1 = False
        RunTest2 = False
        RunTest3 = False
        RunTest4 = True
        RunTest5 = False
        AllManualTests("Tuning")
    End Sub
    Private Sub cmdStartTest5_Click(sender As Object, e As EventArgs) Handles cmdStartTest5.Click
        SaveData = False
        RunTest1 = False
        RunTest2 = False
        RunTest3 = False
        RunTest4 = False
        RunTest5 = True
        AllManualTests("Tuning")
    End Sub

    Private Sub Before_Click(sender As Object, e As EventArgs) Handles Before.Click
        SaveData = True
        RunTest1 = True
        RunTest2 = True
        RunTest3 = True
        RunTest4 = True
        RunTest5 = True
        AllManualTests("Before Tuning")
    End Sub

    Private Sub After_Click(sender As Object, e As EventArgs) Handles After.Click
        SaveData = True
        RunTest1 = True
        RunTest2 = True
        RunTest3 = True
        RunTest4 = True
        RunTest5 = True
        AllManualTests("After Tuning")
    End Sub


    Private Sub SaveStats_Click(sender As Object, e As EventArgs) Handles SaveStats.Click
        Try
            If txtQtyTest.Text <> "0" And txtQtyTuned.Text <> "0" And txtQtyTest.Text <> "" And txtQtyTuned.Text <> "" Then
                Dim KitQty As String = txtKitQty.Text
                Dim QtyTest As String = txtQtyTest.Text
                Dim KitQtyTuned As String = txtQtyTuned.Text
                Dim item1 As String = txtItem1.Text
                Dim part1 As String = txtItem1.Text
                Dim Desc1 As String = txtDesc1.Text
                Dim item2 As String = txtItem2.Text
                Dim part2 As String = txtItem2.Text
                Dim Desc2 As String = txtDesc2.Text
                Dim item3 As String = txtItem3.Text
                Dim part3 As String = txtItem3.Text
                Dim Desc3 As String = txtDesc3.Text
                Dim percent As String = (Math.Round((KitQtyTuned / QtyTest), 2) * 100)
                SaveTuning(KitQty, QtyTest, KitQtyTuned, percent, item1, part1, Desc1, item2, part2, Desc2, item3, part3, Desc3)
            Else
                MYMsgBox("Please fill in all parameters")
            End If
        Catch ex As Exception
            MYMsgBox("Please fill in all parameters correctly")
        End Try
    End Sub

    Private Sub txtKitQty_TextChanged(sender As Object, e As EventArgs) Handles txtKitQty.TextChanged
        Try
            If txtQtyTest.Text <> "0" And txtQtyTuned.Text <> "0" And txtQtyTest.Text <> "" And txtQtyTuned.Text <> "" Then
                Dim QtyTest As String = txtQtyTest.Text
                Dim KitQtyTuned As String = txtQtyTuned.Text
                txtPercent.Text = (Math.Round((KitQtyTuned / QtyTest), 2) * 100) & "%"
            End If
        Catch ex As Exception
            MYMsgBox("Please fill in all parameters correctly")
        End Try
    End Sub

    Private Sub txtQtyTest_TextChanged(sender As Object, e As EventArgs) Handles txtQtyTest.TextChanged
        Try
            If txtQtyTest.Text <> "0" And txtQtyTuned.Text <> "0" And txtQtyTest.Text <> "" And txtQtyTuned.Text <> "" Then
                Dim QtyTest As String = txtQtyTest.Text
                Dim KitQtyTuned As String = txtQtyTuned.Text
                txtPercent.Text = (Math.Round((KitQtyTuned / QtyTest), 2) * 100) & "%"
            End If
        Catch ex As Exception
            MYMsgBox("Please fill in all parameters correctly")
        End Try
    End Sub

    Private Sub txtQtyTuned_TextChanged(sender As Object, e As EventArgs) Handles txtQtyTuned.TextChanged
        Try
            If txtQtyTest.Text <> "0" And txtQtyTuned.Text <> "0" And txtQtyTest.Text <> "" And txtQtyTuned.Text <> "" Then
                Dim QtyTest As String = txtQtyTest.Text
                Dim KitQtyTuned As String = txtQtyTuned.Text
                txtPercent.Text = (Math.Round((KitQtyTuned / QtyTest), 2) * 100) & "%"
            End If
        Catch ex As Exception
            MYMsgBox("Please fill in all parameters correctly")
        End Try
    End Sub

    Private Sub btSwitch1_Click(sender As Object, e As EventArgs) Handles btSwitch1.Click
        Dim status As Integer
        Dim StatusRet As Integer
        If cmbSwitch.Text = "Switch POS 1" Then
            btSwitch1.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 2" Then
            btSwitch2.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 3" Then
            btSwitch3.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 4" Then
            btSwitch4.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 5" Then
            btSwitch5.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 6" Then
            btSwitch6.BackColor = Color.Gold
        End If
        btSwitch1.BackColor = Color.Red
        cmbSwitch.Text = "Switch POS 1"
        Me.Refresh()
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
    End Sub

    Private Sub btSwitch2_Click(sender As Object, e As EventArgs) Handles btSwitch2.Click
        Dim status As Integer
        Dim StatusRet As Integer
        If cmbSwitch.Text = "Switch POS 1" Then
            btSwitch1.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 2" Then
            btSwitch2.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 3" Then
            btSwitch3.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 4" Then
            btSwitch4.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 5" Then
            btSwitch5.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 6" Then
            btSwitch6.BackColor = Color.Gold
        End If
        btSwitch2.BackColor = Color.Red
        cmbSwitch.Text = "Switch POS 2"
        Me.Refresh()
        status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
    End Sub

    Private Sub btSwitch3_Click(sender As Object, e As EventArgs) Handles btSwitch3.Click
        Dim status As Integer
        Dim StatusRet As Integer
        If cmbSwitch.Text = "Switch POS 1" Then
            btSwitch1.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 2" Then
            btSwitch2.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 3" Then
            btSwitch3.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 4" Then
            btSwitch4.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 5" Then
            btSwitch5.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 6" Then
            btSwitch6.BackColor = Color.Gold
        End If
        btSwitch3.BackColor = Color.Red
        cmbSwitch.Text = "Switch POS 3"
        Me.Refresh()
        status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
    End Sub

    Private Sub btSwitch4_Click(sender As Object, e As EventArgs) Handles btSwitch4.Click
        Dim status As Integer
        Dim StatusRet As Integer
        If cmbSwitch.Text = "Switch POS 1" Then
            btSwitch1.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 2" Then
            btSwitch2.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 3" Then
            btSwitch3.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 4" Then
            btSwitch4.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 5" Then
            btSwitch5.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 6" Then
            btSwitch6.BackColor = Color.Gold
        End If

        btSwitch4.BackColor = Color.Red
        Me.Refresh()
        cmbSwitch.Text = "Switch POS 4"
        status = SwitchCom.SetSwitchPosition(4) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
    End Sub
    Private Sub btSwitch5_Click(sender As Object, e As EventArgs) Handles btSwitch5.Click
        Dim status As Integer
        Dim StatusRet As Integer
        If cmbSwitch.Text = "Switch POS 1" Then
            btSwitch1.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 2" Then
            btSwitch2.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 3" Then
            btSwitch3.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 4" Then
            btSwitch4.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 5" Then
            btSwitch5.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 6" Then
            btSwitch6.BackColor = Color.Gold
        End If
        btSwitch5.BackColor = Color.Red
        Me.Refresh()
        cmbSwitch.Text = "Switch POS 5"
        status = SwitchCom.SetSwitchPosition(5) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
    End Sub
    Private Sub btSwitch6_Click(sender As Object, e As EventArgs) Handles btSwitch6.Click
        Dim status As Integer
        Dim StatusRet As Integer
        If cmbSwitch.Text = "Switch POS 1" Then
            btSwitch1.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 2" Then
            btSwitch2.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 3" Then
            btSwitch3.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 4" Then
            btSwitch4.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 5" Then
            btSwitch5.BackColor = Color.Gold
        ElseIf cmbSwitch.Text = "Switch POS 6" Then
            btSwitch6.BackColor = Color.Gold
        End If
        btSwitch6.BackColor = Color.Red
        Me.Refresh()
        cmbSwitch.Text = "Switch POS 6"
        status = SwitchCom.SetSwitchPosition(6) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
        ' StatusLog.Items.Add("Switch POS 4:" & "" & DateTime.Now.ToString)
    End Sub
    Private Sub btSpecs_Click(sender As Object, e As EventArgs) Handles btSpecs.Click
        Dim SPEC As New frmSpecifications
        Me.Hide()
        ActivePage = "Manual"
        SPEC.StartPosition = FormStartPosition.Manual
        SPEC.Location = New Point(globals.XLocation, globals.YLocation)
        SPEC.ShowDialog()
        LoadSpecs()
        Me.Show()
    End Sub
    Private Sub txtArtwork_TextChanged(sender As Object, e As EventArgs) Handles txtArtwork.TextChanged
        UUTReset = True
        txtArtwork.SelectionStart = Len(txtArtwork.Text)
        txtArtwork.Text = Trim(txtArtwork.Text.ToUpper)
        If txtLOT.Text.Length() = 1 Then
            txtLOT.Text = "000000000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 2 Then
            txtLOT.Text = "00000000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 3 Then
            txtLOT.Text = "0000000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 4 Then
            txtLOT.Text = "000000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 5 Then
            txtLOT.Text = "00000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 6 Then
            txtLOT.Text = "0000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 7 Then
            txtLOT.Text = "000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 8 Then
            txtLOT.Text = "00000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 9 Then
            txtLOT.Text = "0000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 10 Then
            txtLOT.Text = "000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 11 Then
            txtLOT.Text = "00" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 12 Then
            txtLOT.Text = "0" + txtLOT.Text
        End If
        ArtworkRevision = Artwork + Rev + Panel + Sector + LOT
       
    End Sub

    Private Sub txtPanel_TextChanged(sender As Object, e As EventArgs) Handles txtPanel.TextChanged
        UUTReset = True
        txtPanel.SelectionStart = Len(txtPanel.Text)
        txtPanel.Text = Trim(txtPanel.Text.ToUpper)
        If txtPanel.Text.Length() = 1 Then
            If txtPanel.Text.Contains("*") Then
                txtPanel.Text = "*" + txtPanel.Text
            Else
                txtPanel.Text = "0" + txtPanel.Text
            End If
        End If
        Panel = txtPanel.Text
        If txtArtwork.Text <> "" And txtSector.Text <> "" And txtLOT.Text <> "" Then
            Artwork = txtArtwork.Text(0)
            If txtArtwork.Text.Length() > 0 Then
                Rev = txtArtwork.Text.Substring(1, 2)
            End If
            Sector = txtSector.Text
            LOT = txtLOT.Text
            ArtworkRevision = Artwork + Rev + Panel + Sector + LOT
        End If


    End Sub
    Private Sub txtSector_TextChanged(sender As Object, e As EventArgs) Handles txtSector.TextChanged
        UUTReset = True
        txtSector.SelectionStart = Len(txtSector.Text)
        txtSector.Text = Trim(txtSector.Text.ToUpper)
        Sector = txtSector.Text
        If txtArtwork.Text <> "" And txtPanel.Text <> "" And txtLOT.Text <> "" Then
            Artwork = txtArtwork.Text(0)
            If txtArtwork.Text.Length() > 0 Then
                Rev = txtArtwork.Text.Substring(1, 2)
            End If
            Panel = txtPanel.Text
            LOT = txtLOT.Text
            ArtworkRevision = Artwork + Rev + Panel + Sector + LOT
        End If


    End Sub

    Private Sub txtLOT_TextChanged(sender As Object, e As EventArgs) Handles txtLOT.TextChanged
        UUTReset = True
        txtLOT.SelectionStart = Len(txtSector.Text)
        txtLOT.Text = Trim(txtLOT.Text.ToUpper)
        LOT = txtLOT.Text
        If txtArtwork.Text <> "" And txtPanel.Text <> "" And txtArtwork.Text <> "" Then
            Artwork = txtArtwork.Text(0)
            If txtArtwork.Text.Length() > 0 Then
                Rev = txtArtwork.Text.Substring(1, 2)
            End If
            Panel = txtPanel.Text
            Rev = txtArtwork.Text
            ArtworkRevision = Artwork + Rev + Panel + Sector + LOT
        End If
    End Sub
End Class