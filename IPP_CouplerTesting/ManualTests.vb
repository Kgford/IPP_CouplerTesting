Module ManualTests


    Public Function CalibrateVNA() As Boolean

        Dim status As Long
        Dim StatusRet As Integer
        Dim StartFreq As String
        Dim StopFreq As String
        Dim POW As String

        CalibrateVNA = True
        POW = 0
        If Debug Then
            CalibrateVNA = False
            Exit Function
        End If
        If VNAStr <> "cmbVNA" Then
            SwitchCom.Connect()
            If Not frmMANUALTEST.Simulation.Checked And Not Connected Then
                ScanGPIB.connect("GPIB0::16::INSTR", GetTimeout())
                Connected = True
            End If
            If SwitchedChecked Then  'Auto RF Switching
                status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
                System.Threading.Thread.Sleep(500)
            Else
                MsgBox("Move RF Cables to Position 1", vbOKOnly, "Manual Switch")
            End If
            VNAStr = ScanGPIB.GetModel
            VNAStr = VNAStr

        End If
        frmMANUALTEST.VNAClicked()
        If VNAStr = "HP_8720ES" Then  '20G version
            StartFreq = "50MHZ"
            StopFreq = "20GHZ"
            SetVNAFreq(20000)
            POW = "5"
        ElseIf VNAStr = "AG_E5071B" Then
            StartFreq = "300E3"
            StopFreq = "8500E6"
            SetVNAFreq(85000)
            POW = "5"
        ElseIf VNAStr = "N3383A" Then
            StartFreq = "300E3"
            StopFreq = "9000E6"
            SetVNAFreq(90000)
            POW = "5"
        ElseIf VNAStr = "HP_8753E" Then
            StartFreq = "10KHZ"
            StopFreq = "6GHZ"
            SetVNAFreq(6000)
            POW = "10"
        ElseIf VNAStr = "HP_8753C" Then
            ScanGPIB.BusWrite("PRES;")
            If MsgBox("Calibrate to 6GHZ ?", vbYesNo, "High Frequency") = vbYes Then
                MsgBox("Turn Doubler on ", , "System - Freq Range 6GHz")
                StartFreq = "3MHZ"
                StopFreq = "6GHZ"
                SetVNAFreq(6000)
                POW = "20"
            Else
                MsgBox("Set VNA to Low Frequency Mode")
                StartFreq = "30KHZ"
                StopFreq = "3GHZ"
                SetVNAFreq(3000)
                POW = "20"
            End If
        Else
            StartFreq = "30KHZ"
            StopFreq = "3GHZ"
            SetVNAFreq(3000)
        End If

        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":SYST:PRES;")
            System.Threading.Thread.Sleep(2000)
            ScanGPIB.BusWrite("SENS:SWE:POIN 1601")
            ScanGPIB.BusWrite("SENS:FREQ:STAR " & StartFreq)
            ScanGPIB.BusWrite("SENS:FREQ:STOP " & StopFreq)

            ScanGPIB.BusWrite(":SOUR1:POW " & POW)

            ScanGPIB.BusWrite(":DISP:WIND1:TRAC1:STAT ON")
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON")

            ScanGPIB.BusWrite(":DISP:SPL D1")
            ScanGPIB.BusWrite(":DISP:WIND1:SPL D1_2")
            ScanGPIB.BusWrite(":CALC1:PAR:COUN 2")
            ScanGPIB.BusWrite(":CALC1:PAR1:DEF S11")
            ScanGPIB.BusWrite(":CALC1:PAR1:SEL")
            ScanGPIB.BusWrite(":CALC1:FORM MLOG")

            ScanGPIB.BusWrite(":DISP:WIND1:TRAC1:Y:RLEV 0")
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC1:Y:PDIV 10")

            ScanGPIB.BusWrite(":CALC1:PAR2:DEF S21")
            ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
            ScanGPIB.BusWrite(":CALC1:FORM MLOG")
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())

            Test = GetSpecification("AmplitudeBalance")
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
            ScanGPIB.BusWrite(":TRIG:SOUR INT")

        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":SYST:PRES;")
            System.Threading.Thread.Sleep(2000)
            ScanGPIB.BusWrite("CALC:PAR:DEL:ALL")
            ScanGPIB.BusWrite("DISP:TILE")
            ScanGPIB.BusWrite("DISP:WIND2 ON")

            ScanGPIB.BusWrite("DISP:WIND1:ENAB ON")
            ScanGPIB.BusWrite("DISP:WIND2:ENAB ON")

            ScanGPIB.BusWrite("CALC1:PAR:DEF 'CH1_S11_1',S11")
            ScanGPIB.BusWrite("DISP:WIND1:TRAC1:FEED 'CH1_S11_1'")
            ScanGPIB.BusWrite("CALC1:FORM MLOG")

            ScanGPIB.BusWrite("DISP:WIND1:TRAC1:Y:RLEV 0")
            ScanGPIB.BusWrite("DISP:WIND1:TRAC1:Y:PDIV 10")

            ScanGPIB.BusWrite("CALC1:PAR1:DEF 'CH1_S21_1',S21")
            ScanGPIB.BusWrite("DISP:WIND2:TRAC2:FEED 'CH1_S21_1'")
            ScanGPIB.BusWrite("CALC1:FORM MLOG")

            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
            ScanGPIB.BusWrite(":TRIG:SOUR IMM")

            ScanGPIB.BusWrite("SENS1:SWE:POIN 1601")
            ScanGPIB.BusWrite("SENS:FREQ:STAR " & StartFreq)
            ScanGPIB.BusWrite("SENS:FREQ:STOP " & StopFreq)
            ScanGPIB.BusWrite(":SOUR1:POW " & POW)

            Test = GetSpecification("AmplitudeBalance")
        Else
            ScanGPIB.BusWrite("OPC?;CHAN1;")
            ScanGPIB.BusWrite("OPC?;S11;")
            ScanGPIB.BusWrite("OPC?;POWE" & POW & ";")
            ScanGPIB.BusWrite("OPC?;LOGM;")
            ScanGPIB.BusWrite("OPC?;SCAL 10")


            ScanGPIB.BusWrite("OPC?;CHAN2;")
            ScanGPIB.BusWrite("OPC?;S21;")
            ScanGPIB.BusWrite("OPC?;LOGM;")
            ScanGPIB.BusWrite("OPC?;SCAL 10")


            ScanGPIB.BusWrite("DUACON;")
            If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
            'If VNAStr = "HP_8753C" Then ScanGPIB.BusWrite("OPC?;SPLID2;")
            ScanGPIB.BusWrite("OPC?;STAR " & StartFreq & ";")
            If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
            ScanGPIB.BusWrite("OPC?;STOP " & StopFreq & ";")
            ScanGPIB.BusWrite("OPC?;POIN1601;")
        End If
        ExtraAvg(2)
        If MsgBox("Perform Full (1P or 2P)Calibration on Switch Poistion 1 then Press OK when complete", vbOKCancel, "Position 1 Full Calibration") = vbCancel Then
            CalibrateVNA = False
            Exit Function
        End If
        CalibrateVNA = True
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":MMEM:STOR 'State01.sta'")
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":MMEM:STOR 'IPP1.cst'")
        Else
            ScanGPIB.BusWrite("SAVE1;")
            ' If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
        End If
        ' If VNAStr = "HP_8753C" Then ScanGPIB.SaveCalCoeff(1)
        If SwitchedChecked Then  'Auto RF Switching
            status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            frmMANUALTEST.cmbSwitch.Text = "POS 2"
            System.Threading.Thread.Sleep(500)
        Else
            MsgBox("Move RF Cables to Position 2", vbOKOnly, "Manual Switch")
        End If
        MsgBox("Calibrate Response Only on Switch Poistion 2 then Press OK when complete", vbOKOnly, "Position 2 Response Calibration Only")
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":MMEM:STOR 'State02.sta'")
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":MMEM:STOR 'IPP2.cst'")
        Else
            ScanGPIB.BusWrite("SAVE2;")
        End If
        ' If VNAStr = "HP_8753C" Then ScanGPIB.SaveCalCoeff(2)

        If SwitchedChecked Then  'Auto RF Switching
            status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            frmMANUALTEST.cmbSwitch.Text = "POS 3"
            System.Threading.Thread.Sleep(500)
        Else
            MsgBox("Move RF Cables to Position 3", vbOKOnly, "Manual Switch")
        End If
        MsgBox("Calibrate Response Only on Switch Poistion 3 then Press OK when complete", vbOKOnly, "Position 3 Response Calibration Only")
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":MMEM:STOR 'State03.sta'")
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":MMEM:STOR 'IPP3.cst'")
        Else
            ScanGPIB.BusWrite("SAVE3;")
        End If
        ' If VNAStr = "HP_8753C" Then ScanGPIB.SaveCalCoeff(3)

        If SwitchedChecked Then  'Auto RF Switching
            status = SwitchCom.SetSwitchPosition(4) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            frmMANUALTEST.cmbSwitch.Text = "POS 4"
            System.Threading.Thread.Sleep(500)
        Else
            MsgBox("Move RF Cables to Position 4", vbOKOnly, "Manual Switch")
        End If
        MsgBox("Calibrate Response Only on Switch Poistion 4 then Press OK when complete", vbOKOnly, "Position 4 Response Calibration Only")
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":MMEM:STOR 'State04.sta'")
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":MMEM:STOR 'IPP4.cst'")
        Else
            ScanGPIB.BusWrite("SAVE4;")
        End If
        'If VNAStr = "HP_8753C" Then ScanGPIB.SaveCalCoeff(4)


        If SwitchedChecked Then  'Auto RF Switching
            status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
            frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
            System.Threading.Thread.Sleep(500)
        Else
            MsgBox("Move RF Cables to Position 1", vbOKOnly, "Manual Switch")
        End If
        CalibrateVNA = True
    End Function


    Public Function InsertionLoss3dB(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim TraceID2 As Long
        Dim Trace1Freq(Points) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim IL1 As Double
        Dim IL2 As Double
        Dim i As Integer
        Dim ILArray(Points) As Double
        Dim ABArray(Points) As Double
        Dim Workstation As String
        Dim Title As String

        InsertionLoss3dB = ""
        Title = ActiveTitle
        ILSetDone = False
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset1.Text = "" Then frmMANUALTEST.txtOffset1.Text = 0
        Spec = GetSpecification("InsertionLoss")
        SwitchCom.Connect()
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset1.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                InsertionLoss3dB = "Pass"
            Else
                InsertionLoss3dB = "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4262
                TraceID2 = 4263
                GetTracePoints(TraceID1)
                Trace1Freq = XArray
                Trace1Data = YArray
                GetTracePoints(TraceID2)
                Trace2Freq = XArray
                Trace2Data = YArray
                Pts = Points
                For i = 0 To Pts
                    IL1 = 1 / (10 ^ (Math.Abs(Trace1Data(i)) * 0.1))
                    IL2 = 1 / (10 ^ (Math.Abs(Trace2Data(i)) * 0.1))
                    ILArray(i) = 10 * Log10(IL1 + IL2)
                Next

                IL1 = MaxNoZero(ILArray)
                IL2 = ILArray.Min

                IL1 = TruncateDecimal(IL1, 2)
                If Right(IL1, 1) = "." Then IL1 = "0" & IL1
                IL2 = TruncateDecimal(IL2, 2)
                If Right(IL2, 1) = "." Then IL2 = "0" & IL2

                If IL1 < IL2 Then
                    IL = IL1
                Else
                    IL = IL2
                End If
                If SpecType.Contains("BALUN") Or Part.Contains("IT") Then
                    IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text) - 3
                Else
                    IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
                End If
                IL = TruncateDecimal(IL, 2)

                If IL <= Spec Then
                    InsertionLoss3dB = "Pass"
                Else
                    InsertionLoss3dB = "Fail"
                End If
            ElseIf PassChecked Then
                IL = Spec
                InsertionLoss3dB = "Pass"
            ElseIf FailChecked Then
                IL = Spec + 10
                InsertionLoss3dB = "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION 1      "
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 1 Then
                    SetSwitchesPosition = 1
                    status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(0)
            Else
                MsgBox("Move Cables to RF Position 1", vbOKOnly, "Manual Switch")
            End If

            If MutiCalChecked Then
                SetupVNA(True, 1)
            End If

            If Not ILSetDone Then
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                Else
                    If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
                    ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("AmplitudeBalance"))
                End If
            End If
            ExtraAvg(2)
            If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
            If VNAStr = "AG_E5071B" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            If VNAStr = "N3383A" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            If VNAStr <> "AG_E5071B" And VNAStr <> "N3383A" Then ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory

            frmMANUALTEST.Refresh()
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & " Insertion Loss J3"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID1 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID1
            End If
            ReDim Preserve IL1Data(Pts)
            ReDim Preserve IL2Data(Pts)
            ScanGPIB.GetTrace(Trace1Freq, IL1Data)
            Trace1Freq = TrimX(Trace1Freq)
            IL1Data_offs = TrimY(IL1Data, CDbl(frmMANUALTEST.txtOffset1.Text))
            If TraceChecked And Not TweakMode Then
                If UUTNum_Reset <= 5 Then
                    ReDim Preserve XArray(IL1Data.Count - 1)
                    ReDim Preserve YArray(IL1Data.Count - 1)
                    Array.Clear(YArray, 0, IL1Data.Count - 1)
                    XArray = Trace1Freq
                    YArray = IL1Data
                    SQL.SaveTrace(Title, TestID, TraceID)
                    YArray = IL1Data_offs
                End If
            End If
            If MutiCalChecked Then ScanGPIB.GetTraceMem()
            ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION 2       "
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 2 Then
                    SetSwitchesPosition = 2
                    status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Bi
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(1)
            Else
                SetSwitchesPosition = 2
                MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
            End If

            If MutiCalChecked Then
                SetupVNA(True, 2)
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                    ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                    ScanGPIB.BusWrite(":CALC:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(":CALC:MATH:MEM") 'Data into Memory
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
                    ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("AmplitudeBalance"))
                    ScanGPIB.BusWrite("OPC?;INPUDATA," & gBuffer & ";") ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
                End If
            End If
            ExtraAvg(2)
            frmMANUALTEST.Refresh()
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":INIT:CONT ON")  'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM ON") 'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":INIT:CONT ON")  'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM ON") 'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON")  ' Data On
            Else
                ScanGPIB.BusWrite("OPC?;CONT") ' and start another sweep
                ScanGPIB.BusWrite("DISPDATM;")  'Data and Memory
            End If

            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Manual " & TestRun & " Insertion Loss J4"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID2 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID2
            End If
            ScanGPIB.GetTrace(Trace2Freq, IL2Data)
            Trace2Freq = TrimX(Trace2Freq)
            IL2Data_offs = TrimY(IL2Data, CDbl(frmMANUALTEST.txtOffset1.Text))
            If TraceChecked And Not TweakMode Then
                ReDim Preserve XArray(IL2Data.Count - 1)
                ReDim Preserve YArray(IL2Data.Count - 1)
                Array.Clear(YArray, 0, IL2Data.Count - 1)
                XArray = Trace2Freq
                YArray = IL2Data
                SQL.SaveTrace(Title, TestID, TraceID)
                YArray = IL2Data_offs
                If UUTNum_Reset <= 5 Then
                    For x = 0 To YArray.Count - 1
                        IL2_YArray(UUTNum_Reset - 1, x) = YArray(x)
                    Next
                End If
            End If

            'TraceID1 =265
            'TraceID2 =266
            Pts = Points
            'If TraceChecked And Not TweakMode Then  ' Simulated Trace Data
            '    GetTracePoints(TraceID1)
            '    IL1Data = YArray
            '    GetTracePoints(TraceID2)
            '    IL2Data = YArray
            'End If

            For i = 0 To Pts
                ReDim Preserve ILArray(i)
                IL1 = 1 / (10 ^ (Math.Abs(IL1Data(i)) * 0.1))
                IL2 = 1 / (10 ^ (Math.Abs(IL2Data(i)) * 0.1))
                ILArray(i) = 10 * Log10(IL1 + IL2)
            Next

            'IL1 = MaxNoZero(ILArray)
            'IL2 = MinNoZero(ILArray)


            'IL1 = Math.Round(IL1, 3)
            'If Right(IL1, 1) = "." Then IL1 = "0" & IL1
            'IL2 = Math.Round(IL2, 3)
            'If Right(IL2, 1) = "." Then IL2 = "0" & IL2

            ' If IL1 < IL2 Then
            'IL = IL1
            'Else
            ' IL = IL2
            ' End If

            IL = ILArray(Pts - 1)
            If SpecType.Contains("BALUN") Or Part.Contains("IT") Then
                IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text) - 3
            Else
                IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
            End If
            IL = TruncateDecimal(IL, 3)


            If IL <= Spec Then
                InsertionLoss3dB = "Pass"
            Else
                InsertionLoss3dB = "Fail"
            End If

            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON")  ' Data On
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF")  'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
            Else
                ScanGPIB.BusWrite("OPC?;DISPDATA")
            End If
        End If
        ABTraceID1 = TraceID1
        ABTraceID2 = TraceID2
        frmMANUALTEST.Refresh()
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.Refresh()
    End Function
    Public Function InsertionLoss3dB_multiband(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim TraceID2 As Long
        Dim Trace1Freq(Points) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim Trace1Freq1(Points) As Double
        Dim Trace1Data1(Points) As Double
        Dim Trace2Freq1(Points) As Double
        Dim Trace2Data1(Points) As Double
        Dim Trace1Freq2(Points) As Double
        Dim Trace1Data2(Points) As Double
        Dim Trace2Freq2(Points) As Double
        Dim Trace2Data2(Points) As Double
        Dim IL1 As Double
        Dim IL2 As Double
        Dim IL3 As Double
        Dim IL4 As Double
        Dim i As Integer
        Dim ILArray(Points) As Double
        Dim ILArray1(Points) As Double
        Dim ILArray2(Points) As Double
        Dim ABArray(Points) As Double
        Dim Workstation As String
        Dim Title As String

        InsertionLoss3dB_multiband = ""
        Title = ActiveTitle
        ILSetDone = False
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset1.Text = "" Then frmMANUALTEST.txtOffset1.Text = 0
        Spec = GetSpecification("InsertionLoss")
        SwitchCom.Connect()
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset1.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                InsertionLoss3dB_multiband = "Pass"
            Else
                InsertionLoss3dB_multiband = "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4262
                TraceID2 = 4263
                GetTracePoints(TraceID1)
                Trace1Freq = XArray
                Trace1Data = YArray
                GetTracePoints(TraceID2)
                Trace2Freq = XArray
                Trace2Data = YArray
                Pts = Points
                For i = 0 To Pts
                    IL1 = 1 / (10 ^ (Math.Abs(Trace1Data(i)) * 0.1))
                    IL2 = 1 / (10 ^ (Math.Abs(Trace2Data(i)) * 0.1))
                    ILArray(i) = 10 * Log10(IL1 + IL2)
                Next

                IL1 = MaxNoZero(ILArray)
                IL2 = ILArray.Min

                IL1 = TruncateDecimal(IL1, 2)
                If Right(IL1, 1) = "." Then IL1 = "0" & IL1
                IL2 = TruncateDecimal(IL2, 2)
                If Right(IL2, 1) = "." Then IL2 = "0" & IL2

                If IL1 < IL2 Then
                    IL = IL1
                Else
                    IL = IL2
                End If
                IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
                IL = TruncateDecimal(IL, 2)

                If IL <= Spec Then
                    InsertionLoss3dB_multiband = "Pass"
                Else
                    InsertionLoss3dB_multiband = "Fail"
                End If
            ElseIf PassChecked Then
                IL = Spec
                InsertionLoss3dB_multiband = "Pass"
            ElseIf FailChecked Then
                IL = Spec + 10
                InsertionLoss3dB_multiband = "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION 1     Band 1 "
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 1 Then
                    SetSwitchesPosition = 1
                    status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(0)
            Else
                MsgBox("Move Cables to RF Position 1", vbOKOnly, "Manual Switch")
            End If

            If MutiCalChecked Then
                SetupVNA(True, 1)
            End If

            If Not ILSetDone Then
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite("SENS:FREQ:STAR " & SpecAB_start1 & "E6;")
                    ScanGPIB.BusWrite("SENS:FREQ:STOP " & SpecAB_stop1 & "E6;")
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("SENS1:FREQ:STAR " & SpecAB_start1 & "E6;")
                    ScanGPIB.BusWrite("SENS1:FREQ:STOP " & SpecAB_stop1 & "E6;")
                    ScanGPIB.BusWrite("CALC:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                Else
                    If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
                    ScanGPIB.BusWrite("STAR " & SpecAB_start1 & "MHZ;")
                    ScanGPIB.BusWrite("STOP " & SpecAB_stop1 & "MHZ;")
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
                    ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("AmplitudeBalance"))
                End If
            End If
            ExtraAvg(2)
            If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
            If VNAStr = "AG_E5071B" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            If VNAStr = "N3383A" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            If VNAStr <> "AG_E5071B" And VNAStr <> "N3383A" Then ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory

            frmMANUALTEST.Refresh()
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & " Insertion Loss J3"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID1 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID1
            End If
            ReDim Preserve Trace1Freq1(Pts)
            ReDim Preserve IL1Data_offs1(Pts)
            ScanGPIB.GetTrace(Trace1Freq1, IL1Data1, SpecAB_start1, SpecAB_stop1)
            Trace1Freq = TrimX(Trace1Freq)
            IL1Data_offs1 = TrimY(IL1Data1, CDbl(frmMANUALTEST.txtOffset1.Text))
            If TraceChecked And Not TweakMode Then
                If UUTNum_Reset <= 5 Then
                    ReDim Preserve XArray(IL1Data1.Count - 1)
                    ReDim Preserve YArray(IL1Data1.Count - 1)
                    Array.Clear(YArray, 0, IL1Data1.Count - 1)
                    XArray = Trace1Freq
                    YArray = IL1Data1
                    SQL.SaveTrace(Title, TestID, TraceID)
                    YArray = IL1Data_offs
                End If
            End If
            If MutiCalChecked Then ScanGPIB.GetTraceMem()
            ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION 2       "
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 2 Then
                    SetSwitchesPosition = 2
                    status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Bi
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(1)
            Else
                SetSwitchesPosition = 2
                MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
            End If

            If MutiCalChecked Then
                SetupVNA(True, 2)
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                    ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                    ScanGPIB.BusWrite(":CALC:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(":CALC:MATH:MEM") 'Data into Memory
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
                    ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("AmplitudeBalance"))
                    ScanGPIB.BusWrite("OPC?;INPUDATA," & gBuffer & ";") ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
                End If
            End If
            ExtraAvg(2)
            frmMANUALTEST.Refresh()
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":INIT:CONT ON")  'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM ON") 'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":INIT:CONT ON")  'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM ON") 'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON")  ' Data On
            Else
                ScanGPIB.BusWrite("OPC?;CONT") ' and start another sweep
                ScanGPIB.BusWrite("DISPDATM;")  'Data and Memory
            End If
            ExtraAvg()
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & " Insertion Loss J4"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID2 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID2
            End If
            ReDim Preserve Trace2Freq1(Pts)
            ReDim Preserve IL2Data_offs1(Pts)
            ScanGPIB.GetTrace(Trace2Freq1, IL2Data1, SpecAB_start1, SpecAB_stop1)
            Trace2Freq1 = TrimX(Trace2Freq1)
            IL2Data_offs1 = TrimY(IL2Data1, CDbl(frmMANUALTEST.txtOffset1.Text))
            If TraceChecked And Not TweakMode Then
                ReDim Preserve XArray(IL2Data1.Count - 1)
                ReDim Preserve YArray(IL2Data1.Count - 1)
                Array.Clear(YArray, 0, IL2Data1.Count - 1)
                XArray = Trace2Freq1
                YArray = IL2Data1
                SQL.SaveTrace(Title, TestID, TraceID)
                YArray = IL2Data_offs1
                If UUTNum_Reset <= 5 Then
                    For x = 0 To YArray.Count - 1
                        IL2_YArray(UUTNum_Reset - 1, x) = YArray(x)
                    Next
                End If
            End If
            '~~~~~~~~~~~~~~~~~~~~~~~~second band~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON")  ' Data On
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF")  'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
            Else
                ScanGPIB.BusWrite("OPC?;DISPDATA")
            End If
            ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION 1     Band 2 "
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 1 Then
                    SetSwitchesPosition = 1
                    status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(0)
            Else
                MsgBox("Move Cables to RF Position 1", vbOKOnly, "Manual Switch")
            End If

            If MutiCalChecked Then
                SetupVNA(True, 1)
            End If
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite("SENS:FREQ:STAR " & SpecAB_start2 & "E6;")
                ScanGPIB.BusWrite("SENS:FREQ:STOP " & SpecAB_stop2 & "E6;")
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("SENS1:FREQ:STAR " & SpecAB_start2 & "E6;")
                ScanGPIB.BusWrite("SENS1:FREQ:STOP " & SpecAB_stop2 & "E6;")
            Else
                If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
                ScanGPIB.BusWrite("STAR " & SpecAB_start2 & "MHZ;")
                ScanGPIB.BusWrite("STOP " & SpecAB_stop2 & "MHZ;")
            End If
            Delay(3000)
            If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
            If VNAStr = "AG_E5071B" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            If VNAStr = "N3383A" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            If VNAStr <> "AG_E5071B" And VNAStr <> "N3383A" Then ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory

            frmMANUALTEST.Refresh()
            ReDim Preserve Trace1Freq2(Pts)
            ReDim Preserve IL1Data_offs2(Pts)
            ScanGPIB.GetTrace(Trace1Freq2, IL1Data2, SpecAB_start2, SpecAB_stop2)
            Trace1Freq2 = TrimX(Trace1Freq2)
            IL1Data_offs2 = TrimY(IL1Data2, CDbl(frmMANUALTEST.txtOffset1.Text))
            If MutiCalChecked Then ScanGPIB.GetTraceMem()
            ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION 2      Band2 "
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 2 Then
                    SetSwitchesPosition = 2
                    status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Bi
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(1)
            Else
                SetSwitchesPosition = 2
                MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
            End If

            If MutiCalChecked Then
                SetupVNA(True, 2)
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                    ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                    ScanGPIB.BusWrite(":CALC:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(":CALC:MATH:MEM") 'Data into Memory
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
                    ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("AmplitudeBalance"))
                    ScanGPIB.BusWrite("OPC?;INPUDATA," & gBuffer & ";") ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
                End If
            End If
            frmMANUALTEST.Refresh()
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":INIT:CONT ON")  'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM ON") 'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":INIT:CONT ON")  'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM ON") 'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON")  ' Data On
            Else
                ScanGPIB.BusWrite("OPC?;CONT") ' and start another sweep
                ScanGPIB.BusWrite("DISPDATM;")  'Data and Memory
            End If

            ReDim Preserve Trace2Freq2(Pts)
            ReDim Preserve IL2Data_offs2(Pts)
            ScanGPIB.GetTrace(Trace2Freq1, IL2Data2, SpecAB_start2, SpecAB_stop2)
            Trace2Freq2 = TrimX(Trace1Freq2)
            IL2Data_offs2 = TrimY(IL2Data2, CDbl(frmMANUALTEST.txtOffset1.Text))

            Pts = Points

            For i = 0 To Pts - 1
                IL1 = 1 / (10 ^ (Math.Abs(IL1Data1(i)) * 0.1))
                IL2 = 1 / (10 ^ (Math.Abs(IL2Data1(i)) * 0.1))
                ILArray1(i) = 10 * Log10(IL1 + IL2)
            Next

            IL1 = MaxNoZero(ILArray1)
            IL2 = MinNoZero(ILArray1)

            For i = 0 To Pts - 1
                IL3 = 1 / (10 ^ (Math.Abs(IL1Data2(i)) * 0.1))
                IL4 = 1 / (10 ^ (Math.Abs(IL2Data2(i)) * 0.1))
                ILArray2(i) = 10 * Log10(IL3 + IL4)
            Next

            IL3 = MaxNoZero(ILArray1)
            IL4 = MinNoZero(ILArray1)


            IL1 = TruncateDecimal(IL1, 2)
            If Right(IL1, 1) = "." Then IL1 = "0" & IL1
            IL2 = TruncateDecimal(IL2, 2)
            If Right(IL2, 1) = "." Then IL2 = "0" & IL2

            If IL1 < IL2 Then
                IL = IL1
            Else
                IL = IL2
            End If
            IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
            IL = TruncateDecimal(IL, 2)

            frmMANUALTEST.Refresh()
            IL1 = MaxNoZero(ILArray)
            IL1 = TruncateDecimal(IL1, 3)
            If Right(IL1, 1) = "." Then IL1 = "0" & IL1

            IL = Math.Abs(IL1) + CDbl(frmMANUALTEST.txtOffset1.Text)
            IL = TruncateDecimal(IL, 3)

            If IL <= Spec Then
                InsertionLoss3dB_multiband = "Pass"
            Else
                InsertionLoss3dB_multiband = "Fail"
            End If
        End If

        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON")  ' Data On
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF")  'Memory Off"
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
        Else
            ScanGPIB.BusWrite("OPC?;DISPDATA")
            'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
        End If

        ABTraceID1 = TraceID1
        ABTraceID2 = TraceID2
        frmMANUALTEST.Refresh()
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite("SENS:FREQ:STAR " & frmMANUALTEST.txtStartFreq.Text & "E6;")
            ScanGPIB.BusWrite("SENS:FREQ:STOP " & frmMANUALTEST.txtStopFreq.Text & "E6;")
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite("SENS1:FREQ:STAR " & frmMANUALTEST.txtStartFreq.Text & "E6;")
            ScanGPIB.BusWrite("SENS1:FREQ:STOP " & frmMANUALTEST.txtStopFreq.Text & "E6;")
        Else
            If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
            ScanGPIB.BusWrite("STAR " & frmMANUALTEST.txtStartFreq.Text & "MHZ;")
            ScanGPIB.BusWrite("STOP " & frmMANUALTEST.txtStopFreq.Text & "MHZ;")
        End If
    End Function
    Public Function InsertionLoss3dB_marker(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim TraceID2 As Long
        Dim Trace1Freq(Points) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim i As Integer
        Dim ILArray(Points) As Double
        Dim ABArray(Points) As Double
        Dim Workstation As String
        Dim Title As String

        InsertionLoss3dB_marker = ""
        Title = ActiveTitle
        ILSetDone = False
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset1.Text = "" Then frmMANUALTEST.txtOffset1.Text = 0
        Spec = GetSpecification("InsertionLoss")
        SwitchCom.Connect()
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset1.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                InsertionLoss3dB_marker = "Pass"
            Else
                InsertionLoss3dB_marker = "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4262
                TraceID2 = 4263
                GetTracePoints(TraceID1)
                Trace1Freq = XArray
                Trace1Data = YArray
                GetTracePoints(TraceID2)
                Trace2Freq = XArray
                Trace2Data = YArray
                Pts = Points
                For i = 0 To Pts
                    IL1 = 1 / (10 ^ (Math.Abs(Trace1Data(i)) * 0.1))
                    IL2 = 1 / (10 ^ (Math.Abs(Trace2Data(i)) * 0.1))
                    ILArray(i) = 10 * Log10(IL1 + IL2)
                Next

                IL1 = MaxNoZero(ILArray)
                IL2 = ILArray.Min

                IL1 = TruncateDecimal(IL1, 2)
                If Right(IL1, 1) = "." Then IL1 = "0" & IL1
                IL2 = TruncateDecimal(IL2, 2)
                If Right(IL2, 1) = "." Then IL2 = "0" & IL2

                If IL1 < IL2 Then
                    IL = IL1
                Else
                    IL = IL2
                End If

                If SpecType.Contains("BALUN") Or Part.Contains("IT") Then
                    IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text) - 3
                Else
                    IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
                End If
                IL = TruncateDecimal(IL, 2)

                If IL <= Spec Then
                    InsertionLoss3dB_marker = "Pass"
                Else
                    InsertionLoss3dB_marker = "Fail"
                End If
            ElseIf PassChecked Then
                IL = Spec
                InsertionLoss3dB_marker = "Pass"
            ElseIf FailChecked Then
                IL = Spec + 10
                InsertionLoss3dB_marker = "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION 1      "
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 1 Then
                    SetSwitchesPosition = 1
                    status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(1000)
                End If
                frmMANUALTEST.Switch(0)
            Else
                MsgBox("Move Cables to RF Position 1", vbOKOnly, "Manual Switch")
            End If

            If MutiCalChecked Then
                SetupVNA(True, 1)
            End If
            Dim freq = GetSpecification("StopFreqMHz") * 10 ^ 9
            If Not ILSetDone Then
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                    ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                    ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ScanGPIB.BusWrite(":CALC1:MARK1:X " & freq)  'set Marker1 Max freq
                    ExtraAvg()
                    Delay(100)
                    IL1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                    ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                    ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
                    ScanGPIB.BusWrite("CALC:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance") / 10 ^ 3)
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ScanGPIB.BusWrite(":CALC1:MARK1:X " & freq)  'set Marker1 Max freq
                    ExtraAvg(1)
                    If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
                    If VNAStr = "AG_E5071B" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                    If VNAStr = "N3383A" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                    If VNAStr <> "AG_E5071B" And VNAStr <> "N3383A" Then ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
                    IL1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                Else
                    If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
                    ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("AmplitudeBalance"))
                    ScanGPIB.BusWrite("MARK1;")  'Marker 1 on
                    ScanGPIB.BusWrite("MARKBUCK 200;")  'Marker 1 max freq
                    ExtraAvg(2)
                    Delay(1000)
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    IL1 = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
                End If
            End If
            If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
            If VNAStr = "AG_E5071B" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            If VNAStr = "N3383A" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            If VNAStr <> "AG_E5071B" And VNAStr <> "N3383A" Then ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory

            frmMANUALTEST.Refresh()
            ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION 2       "
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 2 Then
                    SetSwitchesPosition = 2
                    status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Bi
                    System.Threading.Thread.Sleep(1000)
                End If
                frmMANUALTEST.Switch(1)
            Else
                SetSwitchesPosition = 2
                MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
            End If
            ExtraAvg(2)
            frmMANUALTEST.Refresh()

            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker1 on
                ScanGPIB.BusWrite(":CALC1:MARK1:X " & freq)  'set Marker1 Max freq
                ExtraAvg()
                Delay(100)
                IL2 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker1 on
                ScanGPIB.BusWrite(":CALC1:MARK1:X " & freq)  'set Marker2 Max freq
                ExtraAvg()
                IL2 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
            Else
                ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
                ScanGPIB.BusWrite("MARK1;")  'Marker 1 on
                ScanGPIB.BusWrite("OPC?;CONT") ' and start another sweep
                ScanGPIB.BusWrite("MARKBUCK 200;")  'Marker 1 max freq
                ExtraAvg(2)
                IL2 = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
            End If
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":INIT:CONT ON")  'Memory On"
                'ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM ON") 'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 0.5")
                ScanGPIB.BusWrite(":CALC1:MARK2 ON")  'Marker2 on
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MIN")  'Marker 2 min
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MIN")  'Marker 2 min
                Delay(100)
                IL1AB = ScanGPIB.MarkerQuery(":CALC1:MARK2:Y?")  'Get Marker2 val
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MAX")  'Marker 2 max
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MAX")  'Marker 2 max
                Delay(100)
                IL2AB = ScanGPIB.MarkerQuery(":CALC1:MARK2:Y?")  'Get Marker2 val
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":INIT:CONT ON")  'Memory On"
                'ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM ON") 'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON")  ' Data On
                ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 0.5")
                ScanGPIB.BusWrite(":CALC1:MARK2 ON")  'Marker2 on
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MIN")  'Marker 2 min
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MIN")  'Marker 2 min
                IL1AB = ScanGPIB.MarkerQuery(":CALC1:MARK2:Y?")  'Get Marker2 val
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MAX")  'Marker 2 max
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MAX")  'Marker 2 max
                IL2AB = ScanGPIB.MarkerQuery(":CALC1:MARK2:Y?")  'Get Marker2 val
            Else
                ScanGPIB.BusWrite("OPC?;CONT") ' and start another sweep
                'ScanGPIB.BusWrite("DISPDATM;")  'Data and Memory
                ScanGPIB.BusWrite("OPC?;DISPDDM;") 'Data/Memory
                ScanGPIB.BusWrite("OPC?;REFV 0")
                ScanGPIB.BusWrite("OPC?;SCAL 0.5")
                ScanGPIB.BusWrite("MARK2;")  'Marker2 on
                ScanGPIB.BusWrite("MARKMINI;")  'Marker 2 min 
                IL1AB = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker2 val
                ScanGPIB.BusWrite("MARKMAXI;")  'Marker 2 max
                IL2AB = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker2 val
            End If
        End If
        frmMANUALTEST.Refresh()

        IL1 = 1 / (10 ^ ((Math.Abs(IL1) * 0.1)))
        IL2 = 1 / (10 ^ ((Math.Abs(IL2) * 0.1)))

        IL1 = TruncateDecimal(IL1, 2)
        If Right(IL1, 1) = "." Then IL1 = "0" & IL1
        IL2 = TruncateDecimal(IL2, 2)
        If Right(IL2, 1) = "." Then IL2 = "0" & IL2

        If IL1 < IL2 Then
            IL = IL1
        Else
            IL = IL2
        End If
        IL = 10 * Log10(IL1 + IL2)
        IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
        If SpecType.Contains("BALUN") Or Part.Contains("IT") Then
            IL = TruncateDecimal(IL, 2) - 3
        Else
            IL = TruncateDecimal(IL, 2)
        End If

        If IL <= Spec Then
            InsertionLoss3dB_marker = "Pass"
        Else
            InsertionLoss3dB_marker = "Fail"
        End If

        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance") / 10 ^ 3)
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON")  ' Data On
            ScanGPIB.BusWrite(":CALC1:MATH:FUNC NORM") 'turn off Math
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance") / 10 ^ 3)
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF")  'Memory Off"
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
            ScanGPIB.BusWrite(":CALC1:MATH:FUNC NORM") 'turn off Math
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
        Else
            ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
            ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("AmplitudeBalance"))
            ScanGPIB.BusWrite("OPC?;DISPDATA")
            ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
        End If
        ABTraceID1 = TraceID1
        ABTraceID2 = TraceID2
        frmMANUALTEST.Refresh()
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.Refresh()
    End Function
    Public Function InsertionLossDIR(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim Trace1Freq(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim IL1 As Double
        Dim i As Integer
        Dim ILArray(Points) As Double
        Dim ABArray(Points) As Double
        Dim Workstation As String
        Dim Title As String

        InsertionLossDIR = ""
        Title = ActiveTitle
        ActiveTitle = "     TESTING INSERTION LOSS     "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset1.Text = "" Then frmMANUALTEST.txtOffset1.Text = 0
        Spec = GetSpecification("InsertionLoss")
        SwitchPorts = SQL.GetSpecification("SwitchPorts")
        SwitchCom.Connect()
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset2.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                InsertionLossDIR = "Pass"
            Else
                InsertionLossDIR = "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4262
                GetTracePoints(TraceID1)

                Pts = Points
                For i = 0 To Pts
                    ILArray(i) = Math.Abs(YArray(i))
                Next

                IL1 = MaxNoZero(ILArray)
                IL1 = TruncateDecimal(IL1, 2)
                If Right(IL1, 1) = "." Then IL1 = "0" & IL1

                IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
                IL = TruncateDecimal(IL, 2)

                If IL <= Spec Then
                    InsertionLossDIR = "Pass"
                Else
                    InsertionLossDIR = "Fail"
                End If
            ElseIf PassChecked Then
                IL = Spec
                InsertionLossDIR = "Pass"
            ElseIf FailChecked Then
                IL = Spec + 10
                InsertionLossDIR = "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION 1      "
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 1 Then
                    SetSwitchesPosition = 1
                    status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(0)
            Else
                MsgBox("Move Cables to RF Position 1", vbOKOnly, "Manual Switch")
            End If

            If MutiCalChecked Then SetupVNA(True, 3)
            frmMANUALTEST.Refresh()

            If Not ILSetDone Then
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV 0")
                    ScanGPIB.BusWrite("OPC?;SCAL 10")

                End If
            End If
            ExtraAvg(2)
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Insertion Loss J3"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID1 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID1
            End If
            ScanGPIB.GetTrace(Trace1Freq, IL1Data)

            Trace1Freq = TrimX(Trace1Freq)
            IL1Data_offs = TrimY(IL1Data, CDbl(frmMANUALTEST.txtOffset1.Text))
            If TraceChecked And Not TweakMode Then
                If UUTNum_Reset <= 5 Then
                    ReDim Preserve XArray(IL1Data.Count - 1)
                    ReDim Preserve YArray(IL1Data.Count - 1)
                    Array.Clear(YArray, 0, IL1Data.Count - 1)
                    XArray = Trace1Freq
                    YArray = IL1Data
                    SQL.SaveTrace(Title, TestID, TraceID)
                    YArray = IL1Data_offs
                End If
            End If
            Pts = Points

            For i = 0 To Pts
                ILArray(i) = IL1Data(i)
            Next

            frmMANUALTEST.Refresh()
            IL1 = MinNoZero(ILArray)
            IL1 = TruncateDecimal(IL1, 2)
            If Right(IL1, 1) = "." Then IL1 = "0" & IL1

            IL = Math.Abs(IL1) + CDbl(frmMANUALTEST.txtOffset1.Text)
            IL = TruncateDecimal(IL, 2)

            If IL <= Spec Then
                InsertionLossDIR = "Pass"
            Else
                InsertionLossDIR = "Fail"
            End If
        End If
        COUPTraceID1 = TraceID1
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
        frmMANUALTEST.Switch(0)
    End Function
    Public Function InsertionLossDIR_Marker(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim Trace1Freq(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim IL1 As Double
        Dim i As Integer
        Dim ILArray(Points) As Double
        Dim ABArray(Points) As Double
        Dim Workstation As String
        Dim Title As String

        InsertionLossDIR_Marker = ""
        Title = ActiveTitle
        ActiveTitle = "     TESTING INSERTION LOSS     "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset1.Text = "" Then frmMANUALTEST.txtOffset1.Text = 0
        Spec = GetSpecification("InsertionLoss")
        SwitchCom.Connect()
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset2.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                InsertionLossDIR_Marker = "Pass"
            Else
                InsertionLossDIR_Marker = "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4262
                GetTracePoints(TraceID1)

                Pts = Points
                For i = 0 To Pts
                    ILArray(i) = Math.Abs(YArray(i))
                Next

                IL1 = MaxNoZero(ILArray)
                IL1 = TruncateDecimal(IL1, 2)
                If Right(IL1, 1) = "." Then IL1 = "0" & IL1

                IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
                IL = TruncateDecimal(IL, 2)

                If IL <= Spec Then
                    InsertionLossDIR_Marker = "Pass"
                Else
                    InsertionLossDIR_Marker = "Fail"
                End If
            ElseIf PassChecked Then
                IL = Spec
                InsertionLossDIR_Marker = "Pass"
            ElseIf FailChecked Then
                IL = Spec + 10
                InsertionLossDIR_Marker = "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION 1      "
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 1 Then
                    SetSwitchesPosition = 1
                    status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(0)
            Else
                MsgBox("Move Cables to RF Position 1", vbOKOnly, "Manual Switch")
            End If

            If MutiCalChecked Then SetupVNA(True, 3)
            frmMANUALTEST.Refresh()

            If Not ILSetDone Then
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ExtraAvg()
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MIN")  'Marker 1 min
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")  'Marker 1 
                    System.Threading.Thread.Sleep(500)
                    IL1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ExtraAvg()
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MIN")  'Marker 1 min
                    IL1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV 0")
                    ScanGPIB.BusWrite("OPC?;SCAL 10")
                    ScanGPIB.BusWrite("MARK1;")  'Marker 1 on
                    ExtraAvg(2)
                    ScanGPIB.BusWrite("MARKMINI;")  'Marker 1 min
                    IL1 = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
                End If
            End If

            frmMANUALTEST.Refresh()
            IL1 = TruncateDecimal(IL1, 2)

            IL = Math.Abs(IL1) + CDbl(frmMANUALTEST.txtOffset1.Text)
            IL = TruncateDecimal(IL, 2)

            If IL <= Spec Then
                InsertionLossDIR_Marker = "Pass"
            Else
                InsertionLossDIR_Marker = "Fail"
            End If
        End If
        COUPTraceID1 = TraceID1
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function
    Public Function InsertionLossTRANS(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim Trace1Freq(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim IL1 As Double
        Dim i As Integer
        Dim ILArray(Points) As Double
        Dim ABArray(Points) As Double
        Dim Workstation As String
        Dim Title As String

        InsertionLossTRANS = ""
        Title = ActiveTitle
        ActiveTitle = "     TESTING INSERTION LOSS     "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset1.Text = "" Then frmMANUALTEST.txtOffset1.Text = 0
        Spec = GetSpecification("InsertionLoss")
        SwitchPorts = SQL.GetSpecification("SwitchPorts")
        SwitchCom.Connect()
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset2.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                InsertionLossTRANS = "Pass"
            Else
                InsertionLossTRANS = "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4262
                GetTracePoints(TraceID1)

                Pts = Points
                For i = 0 To Pts
                    ILArray(i) = Math.Abs(YArray(i))
                Next

                IL1 = MaxNoZero(ILArray)
                IL1 = TruncateDecimal(IL1, 2)
                If Right(IL1, 1) = "." Then IL1 = "0" & IL1

                IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
                IL = TruncateDecimal(IL, 2)

                If IL <= Spec Then
                    InsertionLossTRANS = "Pass"
                Else
                    InsertionLossTRANS = "Fail"
                End If
            ElseIf PassChecked Then
                IL = Spec
                InsertionLossTRANS = "Pass"
            ElseIf FailChecked Then
                IL = Spec + 10
                InsertionLossTRANS = "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            ActiveTitle = "     TESTING INSERTION LOSS    "
            If MutiCalChecked Then SetupVNA(True, 3)
            frmMANUALTEST.Refresh()

            If Not ILSetDone Then
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV 0")
                    ScanGPIB.BusWrite("OPC?;SCAL 10")

                End If
            End If
            ExtraAvg(2)
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Insertion Loss"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID1 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID1
            End If
            ReDim Preserve IL1Data(Pts)
            ScanGPIB.GetTrace(Trace1Freq, IL1Data)

            Trace1Freq = TrimX(Trace1Freq)
            IL1Data_offs = TrimY(IL1Data, CDbl(frmMANUALTEST.txtOffset1.Text))
            If TraceChecked And Not TweakMode Then
                If UUTNum_Reset <= 5 Then
                    ReDim Preserve XArray(IL1Data.Count - 1)
                    ReDim Preserve YArray(IL1Data.Count - 1)
                    Array.Clear(YArray, 0, IL1Data.Count - 1)
                    XArray = Trace1Freq
                    YArray = IL1Data
                    SQL.SaveTrace(Title, TestID, TraceID)
                    YArray = IL1Data_offs
                End If
            End If
            Pts = Points
            For i = 0 To Pts
                ILArray(i) = IL1Data(i)
            Next
            frmMANUALTEST.Refresh()
            IL1 = MinNoZero(ILArray)
            IL1 = TruncateDecimal(IL1, 2)
            If Right(IL1, 1) = "." Then IL1 = "0" & IL1

            IL = Math.Abs(IL1) + CDbl(frmMANUALTEST.txtOffset1.Text)
            IL = TruncateDecimal(IL, 2)

            If IL <= Spec Then
                InsertionLossTRANS = "Pass"
            Else
                InsertionLossTRANS = "Fail"
            End If
        End If
        COUPTraceID1 = TraceID1
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
        frmMANUALTEST.Switch(0)
    End Function
    Public Function InsertionLossTRANS_multiband(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim Spec1 As Double
        Dim Spec2 As Double
        Dim TraceID1 As Long
        Dim TraceID2 As Long
        Dim Trace1Freq(Points) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim Trace1Freq1(Points) As Double
        Dim Trace1Data1(Points) As Double
        Dim Trace2Freq1(Points) As Double
        Dim Trace2Data1(Points) As Double
        Dim Trace1Freq2(Points) As Double
        Dim Trace1Data2(Points) As Double
        Dim Trace2Freq2(Points) As Double
        Dim Trace2Data2(Points) As Double
        Dim i As Integer
        Dim ILArray(Points) As Double
        Dim ILArray1(Points) As Double
        Dim ILArray2(Points) As Double
        Dim ABArray(Points) As Double
        Dim Workstation As String
        Dim Title As String

        InsertionLossTRANS_multiband = ""
        Title = ActiveTitle
        ILSetDone = False
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset1.Text = "" Then frmMANUALTEST.txtOffset1.Text = 0
        Spec1 = GetSpecification("InsertionLoss")
        Spec2 = GetSpecification("InsertionLoss2")

        SwitchCom.Connect()
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset1.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec1 Then
                InsertionLossTRANS_multiband = "Pass"
            Else
                InsertionLossTRANS_multiband = "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4262
                TraceID2 = 4263
                GetTracePoints(TraceID1)
                Trace1Freq = XArray
                Trace1Data = YArray
                GetTracePoints(TraceID2)
                Trace2Freq = XArray
                Trace2Data = YArray
                Pts = Points
                For i = 0 To Pts
                    IL1 = 1 / (10 ^ (Math.Abs(Trace1Data(i)) * 0.1))
                    IL2 = 1 / (10 ^ (Math.Abs(Trace2Data(i)) * 0.1))
                    ILArray(i) = 10 * Log10(IL1 + IL2)
                Next

                IL1 = MaxNoZero(ILArray)
                IL2 = ILArray.Min

                IL1 = TruncateDecimal(IL1, 2)
                If Right(IL1, 1) = "." Then IL1 = "0" & IL1
                IL2 = TruncateDecimal(IL2, 2)
                If Right(IL2, 1) = "." Then IL2 = "0" & IL2

                If IL1 < IL2 Then
                    IL = IL1
                Else
                    IL = IL2
                End If
                IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
                IL = TruncateDecimal(IL, 2)

                If IL <= Spec1 Then
                    InsertionLossTRANS_multiband = "Pass"
                Else
                    InsertionLossTRANS_multiband = "Fail"
                End If
            ElseIf PassChecked Then
                IL = Spec1
                InsertionLossTRANS_multiband = "Pass"
            ElseIf FailChecked Then
                IL = Spec1 + 10
                InsertionLossTRANS_multiband = "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            ActiveTitle = "     TESTING INSERTION LOSS Band 1 "
            If MutiCalChecked Then
                SetupVNA(True, 1)
            End If

            If Not ILSetDone Then
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite("SENS:FREQ:STAR " & SpecIL_start1 & "E6;")
                    ScanGPIB.BusWrite("SENS:FREQ:STOP " & SpecIL_stop1 & "E6;")
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV " & GetSpecification("IL2"))
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("SENS1:FREQ:STAR " & SpecIL_start1 & "E6;")
                    ScanGPIB.BusWrite("SENS1:FREQ:STOP " & SpecIL_stop1 & "E6;")
                    ScanGPIB.BusWrite("CALC:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("insertionloss"))
                Else
                    If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
                    ScanGPIB.BusWrite("STAR " & SpecIL_start1 & "MHZ;")
                    ScanGPIB.BusWrite("STOP " & SpecIL_stop1 & "MHZ;")
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
                    ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("insertionloss"))
                End If
            End If
            ExtraAvg(2)

            frmMANUALTEST.Refresh()
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Insertion Loss J3"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID1 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID1
            End If
            ReDim Preserve IL1Data(Pts)
            ReDim Preserve IL2Data(Pts)
            ScanGPIB.GetTrace(Trace1Freq, IL1Data)
            Trace1Freq = TrimX(Trace1Freq)
            IL1Data_offs = TrimY(IL1Data, CDbl(frmMANUALTEST.txtOffset1.Text))
            If TraceChecked And Not TweakMode Then
                If UUTNum_Reset <= 5 Then
                    ReDim Preserve XArray(IL1Data.Count - 1)
                    ReDim Preserve YArray(IL1Data.Count - 1)
                    Array.Clear(YArray, 0, IL1Data.Count - 1)
                    XArray = Trace1Freq
                    YArray = IL1Data
                    SQL.SaveTrace(Title, TestID, TraceID)
                    YArray = IL1Data_offs
                End If
            End If
            Pts = Points

            For i = 0 To Pts
                ILArray(i) = IL1Data(i)
            Next

            frmMANUALTEST.Refresh()
            IL1 = MinNoZero(ILArray)
            IL1 = TruncateDecimal(IL1, 2)
            If Right(IL1, 1) = "." Then IL1 = "0" & IL1

            IL1 = Math.Abs(IL1) + CDbl(frmMANUALTEST.txtOffset1.Text)
            IL1 = TruncateDecimal(IL1, 2)

            '~~~~~~~~~~~~~~~~~~~~~~~~second band~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON")  ' Data On
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF")  'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
            Else
                ScanGPIB.BusWrite("OPC?;DISPDATA")
            End If
            ActiveTitle = "     TESTING INSERTION LOSS Band 2 "

            If MutiCalChecked Then
                SetupVNA(True, 1)
            End If
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite("SENS:FREQ:STAR " & SpecIL_start2 & "E6;")
                ScanGPIB.BusWrite("SENS:FREQ:STOP " & SpecIL_stop2 & "E6;")
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("SENS1:FREQ:STAR " & SpecIL_start2 & "E6;")
                ScanGPIB.BusWrite("SENS1:FREQ:STOP " & SpecIL_stop2 & "E6;")
            Else
                If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
                ScanGPIB.BusWrite("STAR " & SpecIL_start2 & "MHZ;")
                ScanGPIB.BusWrite("STOP " & SpecIL_stop2 & "MHZ;")
            End If
            Delay(3000)
            frmMANUALTEST.Refresh()
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Insertion Loss J4"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID1 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID1
            End If
            ReDim Preserve IL1Data(Pts)
            ReDim Preserve IL2Data(Pts)
            ScanGPIB.GetTrace(Trace2Freq, IL2Data)
            Trace2Freq = TrimX(Trace2Freq)
            IL2Data_offs = TrimY(IL2Data, CDbl(frmMANUALTEST.txtOffset1.Text))
            If TraceChecked And Not TweakMode Then
                If UUTNum_Reset <= 5 Then
                    ReDim Preserve XArray(IL2Data.Count - 1)
                    ReDim Preserve YArray(IL2Data.Count - 1)
                    Array.Clear(YArray, 0, IL2Data.Count - 1)
                    XArray = Trace2Freq
                    YArray = IL2Data
                    SQL.SaveTrace(Title, TestID, TraceID)
                    YArray = IL2Data_offs
                End If
            End If
            Pts = Points

            For i = 0 To Pts
                ILArray(i) = IL2Data(i)
            Next

            frmMANUALTEST.Refresh()
            IL2 = MinNoZero(ILArray)
            IL2 = TruncateDecimal(IL2, 2)
            If Right(IL2, 1) = "." Then IL2 = "0" & IL2

            IL2 = Math.Abs(IL2) + CDbl(frmMANUALTEST.txtOffset1.Text)
            IL2 = TruncateDecimal(IL2, 2)
            Pts = Points

            If IL1 > Spec1 Then
                IL1_status = "Fail"
            End If
            If IL2 > Spec2 Then
                IL2_status = "Fail"
            End If

            If IL1 <= Spec1 And IL2 <= Spec2 Then
                InsertionLossTRANS_multiband = "Pass"
            Else
                InsertionLossTRANS_multiband = "Fail"
            End If
        End If


        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite("SENS:FREQ:STAR " & frmMANUALTEST.txtStartFreq.Text & "E6;")
            ScanGPIB.BusWrite("SENS:FREQ:STOP " & frmMANUALTEST.txtStopFreq.Text & "E6;")
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite("SENS1:FREQ:STAR " & frmMANUALTEST.txtStartFreq.Text & "E6;")
            ScanGPIB.BusWrite("SENS1:FREQ:STOP " & frmMANUALTEST.txtStopFreq.Text & "E6;")
        Else
            If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(2000)
            ScanGPIB.BusWrite("STAR " & frmMANUALTEST.txtStartFreq.Text & "MHZ;")
            ScanGPIB.BusWrite("STOP " & frmMANUALTEST.txtStopFreq.Text & "MHZ;")
        End If
    End Function
    Public Function InsertionLossTRANS_Marker(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim Trace1Freq(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim IL1 As Double
        Dim i As Integer
        Dim ILArray(Points) As Double
        Dim ABArray(Points) As Double
        Dim Workstation As String
        Dim Title As String

        InsertionLossTRANS_Marker = ""
        Title = ActiveTitle
        ActiveTitle = "     TESTING INSERTION LOSS     "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset1.Text = "" Then frmMANUALTEST.txtOffset1.Text = 0
        Spec = GetSpecification("InsertionLoss")
        SwitchCom.Connect()
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset2.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                InsertionLossTRANS_Marker = "Pass"
            Else
                InsertionLossTRANS_Marker = "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4262
                GetTracePoints(TraceID1)

                Pts = Points
                For i = 0 To Pts
                    ILArray(i) = Math.Abs(YArray(i))
                Next

                IL1 = MaxNoZero(ILArray)
                IL1 = TruncateDecimal(IL1, 2)
                If Right(IL1, 1) = "." Then IL1 = "0" & IL1

                IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
                IL = TruncateDecimal(IL, 2)

                If IL <= Spec Then
                    InsertionLossTRANS_Marker = "Pass"
                Else
                    InsertionLossTRANS_Marker = "Fail"
                End If
            ElseIf PassChecked Then
                IL = Spec
                InsertionLossTRANS_Marker = "Pass"
            ElseIf FailChecked Then
                IL = Spec + 10
                InsertionLossTRANS_Marker = "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            ActiveTitle = "     TESTING INSERTION LOSS     "


            If MutiCalChecked Then SetupVNA(True, 3)
            frmMANUALTEST.Refresh()

            If Not ILSetDone Then
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ExtraAvg()
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MIN")  'Marker 1 min
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")  'Marker 1 
                    System.Threading.Thread.Sleep(500)
                    IL1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ExtraAvg()
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MIN")  'Marker 1 min
                    IL1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV 0")
                    ScanGPIB.BusWrite("OPC?;SCAL 10")
                    ScanGPIB.BusWrite("MARK1;")  'Marker 1 on
                    ExtraAvg(2)
                    ScanGPIB.BusWrite("MARKMINI;")  'Marker 1 min
                    IL1 = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
                End If
            End If

            frmMANUALTEST.Refresh()
            IL1 = TruncateDecimal(IL1, 2)

            IL = Math.Abs(IL1) + CDbl(frmMANUALTEST.txtOffset1.Text)
            IL = TruncateDecimal(IL, 2)

            If IL <= Spec Then
                InsertionLossTRANS_Marker = "Pass"
            Else
                InsertionLossTRANS_Marker = "Fail"
            End If
        End If
        COUPTraceID1 = TraceID1
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function


    Public Function InsertionLossCOMB(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim Trace1Freq(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim MinData(32) As Double
        Dim x As Long
        Dim IL1 As Double
        Dim ILArray(Points) As Double
        Dim ABArray(Points) As Double
        Dim Workstation As String
        Dim Title As String
        Dim NumPorts As Double
        Dim PortNum As Byte
        Dim Ports As Integer
        Dim Nominal As Double

        InsertionLossCOMB = ""
        ILSetDone = True
        Title = ActiveTitle
        ActiveTitle = "     TESTING INSERTION LOSS     "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset1.Text = "" Then frmMANUALTEST.txtOffset1.Text = 0
        Spec = GetSpecification("InsertionLoss")
        SwitchCom.Connect()
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset1.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Return "Pass"
            Else
                Return "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4262
                Pts = Points
                GetTracePoints(TraceID1)
                ILArray = YArray

                IL1 = MaxNoZero(ILArray)
                IL1 = TruncateDecimal(IL1, 2)
                If Right(IL1, 1) = "." Then IL1 = "0" & IL1

                IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
                IL = TruncateDecimal(IL, 2)

                If IL <= Spec Then
                    Return "Pass"
                Else
                    Return "Fail"
                End If
            ElseIf PassChecked Then
                IL = Spec
                Return "Pass"
            ElseIf FailChecked Then
                IL = Spec + 10
                Return "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            NumPorts = GetSpecification("Ports")
            PortNum = CByte(NumPorts)
            Ports = Int(NumPorts)

            For x = 1 To Ports
                PortNum = CByte(x)
                ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION " & x & "      "
                If SwitchedChecked Then  'Auto RF Switching
                    SetSwitchesPosition = PortNum
                    status = SwitchCom.SetSwitchPosition(PortNum) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                    frmMANUALTEST.Switch(x - 1)
                Else
                    MsgBox("Move Cables to RF Position " & x & " ")
                End If

                If MutiCalChecked Then SetupVNA(True, x)

                If Not ILSetDone Then
                    If VNAStr = "AG_E5071B" Then
                        ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                        ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                    ElseIf VNAStr = "N3383A" Then
                        ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                        ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                    Else
                        ScanGPIB.BusWrite("OPC?;CHAN2;")
                        ScanGPIB.BusWrite("OPC?;LOGM;")
                        ScanGPIB.BusWrite("OPC?;REFV 0")
                        ScanGPIB.BusWrite("OPC?;SCAL 10")
                    End If
                    ILSetDone = True
                End If
                ExtraAvg(2)
                If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                    Title = "Manual " & TestRun & "Insertion Loss Port " & x & " "
                    SerialNumber = "UUT" & UUTNum_Reset
                    TestID = TestID
                    CalDate = Now
                    Notes = ""
                    Workstation = GetComputerName()
                    TraceID1 = SQL.GetTraceID(Title, TestID)
                    TraceID = TraceID1
                End If
                ScanGPIB.GetTrace(Trace1Freq, IL1Data)
                Trace1Freq = TrimX(Trace1Freq)
                IL1Data_offs = TrimY(IL1Data, CDbl(frmMANUALTEST.txtOffset1.Text))
                If TraceChecked And Not TweakMode Then
                    If UUTNum_Reset <= 5 Then
                        ReDim Preserve XArray(IL1Data.Count - 1)
                        ReDim Preserve YArray(IL1Data.Count - 1)
                        Array.Clear(YArray, 0, IL1Data.Count - 1)
                        XArray = Trace1Freq
                        YArray = IL1Data
                        SQL.SaveTrace(Title, TestID, TraceID)
                        YArray = IL1Data_offs
                    End If
                End If
                MinData(x - 1) = IL1Data.Min

                frmMANUALTEST.Refresh()
            Next x

            For x = 1 To Ports
                If x = 1 Then
                    IL1 = MinData(x)
                Else
                    If MinData(x) < IL1 Then IL1 = MinData(x - 1)
                End If
            Next x
            Nominal = Math.Abs(10 * Log10(1 / Ports))

            IL1 = Math.Abs(Nominal + IL1)

            If Right(IL1, 1) = "." Then IL1 = "0" & IL1

            IL = Math.Abs(IL1) + CDbl(frmMANUALTEST.txtOffset1.Text)
            IL = TruncateDecimal(IL, 2)

            If IL <= Spec Then
                InsertionLossCOMB = "Pass"
            Else
                InsertionLossCOMB = "Fail"
            End If
        End If
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function
    Public Function InsertionLossCOMB_Marker(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim Trace1Freq(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim MinData(32) As Double
        Dim x As Long
        Dim IL1 As Double
        Dim ILArray(Points) As Double
        Dim ABArray(Points) As Double
        Dim Workstation As String
        Dim Title As String
        Dim NumPorts As Double
        Dim PortNum As Byte
        Dim Ports As Integer
        Dim Nominal As Double

        InsertionLossCOMB_Marker = ""
        ILSetDone = True
        Title = ActiveTitle
        ActiveTitle = "     TESTING INSERTION LOSS     "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset1.Text = "" Then frmMANUALTEST.txtOffset1.Text = 0
        Spec = GetSpecification("InsertionLoss")
        SwitchCom.Connect()
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset1.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Return "Pass"
            Else
                Return "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4262
                Pts = Points
                GetTracePoints(TraceID1)
                ILArray = YArray

                IL1 = MaxNoZero(ILArray)
                IL1 = TruncateDecimal(IL1, 2)
                If Right(IL1, 1) = "." Then IL1 = "0" & IL1

                IL = Math.Abs(IL) + CDbl(frmMANUALTEST.txtOffset1.Text)
                IL = TruncateDecimal(IL, 2)

                If IL <= Spec Then
                    Return "Pass"
                Else
                    Return "Fail"
                End If
            ElseIf PassChecked Then
                IL = Spec
                Return "Pass"
            ElseIf FailChecked Then
                IL = Spec + 10
                Return "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            NumPorts = GetSpecification("Ports")
            PortNum = CByte(NumPorts)
            Ports = Int(NumPorts)

            For x = 1 To Ports
                PortNum = CByte(x)
                ActiveTitle = "     TESTING INSERTION LOSS    SW POSITION " & x & "      "
                If SwitchedChecked Then  'Auto RF Switching
                    SetSwitchesPosition = PortNum
                    status = SwitchCom.SetSwitchPosition(PortNum) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                    frmMANUALTEST.Switch(x - 1)
                    frmMANUALTEST.Switch(x - 1)
                Else
                    MsgBox("Move Cables to RF Position " & x & " ")
                End If

                If MutiCalChecked Then SetupVNA(True, x)

                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ExtraAvg(2)
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MIN")  'Marker 1 min
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")  'Marker 1 
                    IL1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ExtraAvg(2)
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MIN")  'Marker 1 min
                    IL1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    ScanGPIB.BusWrite("OPC?;REFV 0")
                    ScanGPIB.BusWrite("OPC?;SCAL 10")
                    ScanGPIB.BusWrite("MARK1;")  'Marker 1 on
                    ExtraAvg(2)
                    ScanGPIB.BusWrite("CHAN2;")
                    ScanGPIB.BusWrite("MARKMINI;")  'Marker 1 min
                    IL1 = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
                End If
                ILSetDone = True

                MinData(x - 1) = IL1
                frmMANUALTEST.Refresh()
            Next x
            ExtraAvg()
            For x = 1 To Ports
                If x = 1 Then
                    IL1 = MinData(x)
                Else
                    If MinData(x) < IL1 Then IL1 = MinData(x - 1)
                End If
            Next x
            Nominal = Math.Abs(10 * Log10(1 / Ports))

            IL1 = Math.Abs(Nominal + IL1)

            If Right(IL1, 1) = "." Then IL1 = "0" & IL1

            IL = Math.Abs(IL1) + CDbl(frmMANUALTEST.txtOffset1.Text)
            IL = TruncateDecimal(IL, 2)

            If IL <= Spec Then
                InsertionLossCOMB_Marker = "Pass"
            Else
                InsertionLossCOMB_Marker = "Fail"
            End If
        End If
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON")  ' Data On
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF")  'Memory Off"
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
        Else
            ScanGPIB.BusWrite("OPC?;DISPDATA")
            ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
        End If
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function

    Public Function ReturnLoss(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim t1 As Trace
        Dim TraceID As Long
        Dim Workstation As String
        Dim Title As String

        Title = ActiveTitle
        ActiveTitle = "     TESTING RETURN LOSS      "
        ReturnLoss = ""
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset2.Text = "" Then frmMANUALTEST.txtOffset3.Text = 0
        Spec = GetSpecification("VSWR")
        SwitchPorts = SQL.GetSpecification("SwitchPorts")
        t1 = New Trace
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset2.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                ReturnLoss = "Pass"
            Else
                ReturnLoss = "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                TraceID = 4264
                If TraceID > 171666 Then
                    GetTracePoints2(TraceID)
                Else
                    GetTracePoints(TraceID)
                End If
                RL = MaxNoZero(YArray)
                RL = RL + CDbl(frmMANUALTEST.txtOffset2.Text)

                If RL <= Spec Then
                    ReturnLoss = "Pass"
                Else
                    ReturnLoss = "Fail"
                End If
            ElseIf PassChecked Then
                RL = Spec
                ReturnLoss = "Pass"
            ElseIf FailChecked Then
                RL = Spec + 10
                ReturnLoss = "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            If SwitchPorts = 1 Then ' ~~~~~~~ 6 port stuff ~~~~~~~~~
                If SwitchedChecked Then  'Auto RF Switching
                    If SetSwitchesPosition <> 5 Then
                        SetSwitchesPosition = 5
                        status = SwitchCom.SetSwitchPosition(5) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)
                    End If
                    frmMANUALTEST.Switch(4)
                Else
                    MsgBox("Move Cables to RF Position 5", vbOKOnly, "Manual Switch")
                End If
            Else ' ~~~~~~~ 4 port stuff ~~~~~~~~~
                If SwitchedChecked Then  'Auto RF Switching
                    If SetSwitchesPosition <> 1 Then
                        SetSwitchesPosition = 1
                        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)
                    End If
                    frmMANUALTEST.Switch(0)
                Else
                    MsgBox("Move Cables to RF Position 1", vbOKOnly, "Manual Switch")
                End If
            End If


            frmMANUALTEST.Refresh()
            ExtraAvg(1)
            System.Threading.Thread.Sleep(100)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR1:SEL")
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC1:PAR1:SEL 'CH1_S11_1'")
            Else
                ScanGPIB.BusWrite("CHAN1;")
            End If
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Return Loss"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID = SQL.GetTraceID(Title, TestID)
            End If
            ScanGPIB.GetTrace_RL(TraceFreq, TraceData)
            TraceFreq = TrimX(TraceFreq)
            TraceData_offs = TrimY(TraceData, CDbl(frmMANUALTEST.txtOffset2.Text))
            If TraceChecked And Not TweakMode Then
                If UUTNum_Reset <= 5 Then
                    ReDim Preserve XArray(TraceData.Count - 1)
                    ReDim Preserve YArray(TraceData.Count - 1)
                    Array.Clear(YArray, 0, TraceData.Count - 1)
                    XArray = TraceFreq
                    YArray = TraceData
                    SQL.SaveTrace(Title, TestID, TraceID)
                    YArray = TraceData_offs
                End If
            End If
            RL = MaxNoZero(TraceData)
            RL = RL + CDbl(frmMANUALTEST.txtOffset2.Text)
            RL = TruncateDecimal(RL, 1)
            If RL <= Spec Then
                ReturnLoss = "Pass"
            Else
                ReturnLoss = "Fail"
            End If
        End If
        frmMANUALTEST.Refresh()
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function
    Public Function ReturnLoss_Marker(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim t1 As Trace
        Dim TraceID As Long
        Dim Workstation As String
        Dim Title As String

        Title = ActiveTitle
        ActiveTitle = "     TESTING RETURN LOSS      "
        ReturnLoss_Marker = ""
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset2.Text = "" Then frmMANUALTEST.txtOffset3.Text = 0
        Spec = GetSpecification("VSWR")
        t1 = New Trace
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset2.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                ReturnLoss_Marker = "Pass"
            Else
                ReturnLoss_Marker = "Fail"
            End If
        ElseIf Debug Then   ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                TraceID = 4264
                If TraceID > 171666 Then
                    GetTracePoints2(TraceID)
                Else
                    GetTracePoints(TraceID)
                End If
                RL = MaxNoZero(YArray)
                RL = RL + CDbl(frmMANUALTEST.txtOffset2.Text)

                If RL <= Spec Then
                    ReturnLoss_Marker = "Pass"
                Else
                    ReturnLoss_Marker = "Fail"
                End If
            ElseIf PassChecked Then
                RL = Spec
                ReturnLoss_Marker = "Pass"
            ElseIf FailChecked Then
                RL = Spec + 10
                ReturnLoss_Marker = "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 1 Then
                    SetSwitchesPosition = 1
                    status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(0)
            Else
                MsgBox("Move Cables to RF Position 1", vbOKOnly, "Manual Switch")
            End If

            frmMANUALTEST.Refresh()
            System.Threading.Thread.Sleep(100)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR1:SEL")
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                ExtraAvg()
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 min
                RL = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC1:PAR1:SEL 'CH1_S11_1'")
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                ExtraAvg()
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                ScanGPIB.BusWrite("CALC1:PAR1:SEL 'CH1_S11_1'")
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 min
                RL = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
            Else
                ScanGPIB.BusWrite("CHAN1;")
                ScanGPIB.BusWrite("OPC?;SCAL 10")
                ScanGPIB.BusWrite("MARK1;")  'Marker 1 on
                ScanGPIB.BusWrite("CHAN1;")
                ScanGPIB.BusWrite("MARKMAXI;")  'Marker 1 min
                ExtraAvg(1)
                RL = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
            End If

            RL = RL + CDbl(frmMANUALTEST.txtOffset2.Text)
            RL = TruncateDecimal(RL, 1)
            If RL <= Spec Then
                ReturnLoss_Marker = "Pass"
            Else
                ReturnLoss_Marker = "Fail"
            End If
            'Turn off all markers
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
            Else
                ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
            End If
        End If
        frmMANUALTEST.Refresh()
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function

    Public Function Isolation(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim TraceID1 As Long
        Dim t1 As Trace
        Dim T1Min As Double
        Dim T1Max As Double
        Dim Workstation As String
        Dim Title As String
        Dim freq As Double

        Title = ActiveTitle
        ActiveTitle = "     TESTING ISOLATION       "
        Isolation = ""

        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset3.Text = "" Then frmMANUALTEST.txtOffset3.Text = 0 ' bad user protection
        t1 = New Trace
        Spec = 0 - GetSpecification("Isolation")
        SwitchPorts = SQL.GetSpecification("SwitchPorts")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset3.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Isolation = "Pass"
            Else
                Isolation = "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                TraceID1 = 4265
                GetTracePoints(TraceID1)
                T1Min = MinNoZero(YArray)
                T1Max = MaxNoZero(YArray)

                ISo = T1Max
                ISo = ISo + CDbl(frmMANUALTEST.txtOffset3.Text)

                If ISo < Spec Then
                    Isolation = "Pass"
                Else
                    Isolation = "Fail"
                End If
            ElseIf PassChecked Then
                ISo = Spec
                Isolation = "Pass"
            ElseIf FailChecked Then
                ISo = SpecISO + 10
                Isolation = "Fail"
            End If
            System.Threading.Thread.Sleep(1000)
            frmMANUALTEST.Refresh()
        Else
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 3 Then
                    SetSwitchesPosition = 3
                    status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(2)
            Else
                MsgBox("Move Cables to RF Position 3", vbOKOnly, "Manual Switch")
            End If
            If MutiCalChecked Then SetupVNA(True, 3)

            System.Threading.Thread.Sleep(100)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & 0 - GetSpecification("Isolation"))
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 2")
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":CALC:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & 0 - GetSpecification("Isolation"))
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 2")
            Else
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("OPC?;LOGM;")
                ScanGPIB.BusWrite("OPC?;REFV " & 0 - GetSpecification("Isolation"))
                ScanGPIB.BusWrite("OPC?;SCAL 2")
            End If
            System.Threading.Thread.Sleep(3000)
            ExtraAvg(2)
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Isolation"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID1 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID1
            End If
            ScanGPIB.GetTrace(TraceFreq, TraceData)
            TraceFreq = TrimX(TraceFreq)
            TraceData_offs = TrimY(TraceData, CDbl(frmMANUALTEST.txtOffset3.Text))
            If TraceChecked And Not TweakMode Then
                If UUTNum_Reset <= 5 Then
                    ReDim Preserve XArray(TraceData.Count - 1)
                    ReDim Preserve YArray(TraceData.Count - 1)
                    Array.Clear(YArray, 0, TraceData.Count - 1)
                    XArray = TraceFreq
                    YArray = TraceData_offs
                    SQL.SaveTrace(Title, TestID, TraceID)
                    YArray = TraceData
                End If
            End If

            frmMANUALTEST.Refresh()
            ISo = MaxNoZero(TraceData)
            freq = CDbl(TraceFreq(MaxX(TraceData)))

            ISo = ISo + CDbl(frmMANUALTEST.txtOffset3.Text)
            frmMANUALTEST.Refresh()
            Spec = 0 - GetSpecification("Isolation", freq)
            'ISo = Format(ISo, 0.00)
            If ISo <= Spec Then
                Isolation = "Pass"
            Else
                Isolation = "Fail"
            End If
        End If
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function
    Public Function Isolation_Marker(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim TraceID1 As Long
        Dim t1 As Trace
        Dim T1Min As Double
        Dim T1Max As Double
        Dim Workstation As String
        Dim Title As String
        Dim freq As Double

        Title = ActiveTitle
        ActiveTitle = "     TESTING ISOLATION       "
        Isolation_Marker = ""

        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset3.Text = "" Then frmMANUALTEST.txtOffset3.Text = 0 ' bad user protection
        t1 = New Trace
        Spec = 0 - GetSpecification("Isolation")
        SwitchPorts = SQL.GetSpecification("SwitchPorts")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset3.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Isolation_Marker = "Pass"
            Else
                Isolation_Marker = "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                TraceID1 = 4265
                GetTracePoints(TraceID1)
                T1Min = MinNoZero(YArray)
                T1Max = MaxNoZero(YArray)

                ISo = T1Max
                ISo = ISo + CDbl(frmMANUALTEST.txtOffset3.Text)

                If ISo < Spec Then
                    Isolation_Marker = "Pass"
                Else
                    Isolation_Marker = "Fail"
                End If
            ElseIf PassChecked Then
                ISo = Spec
                Isolation_Marker = "Pass"
            ElseIf FailChecked Then
                ISo = SpecISO + 10
                Isolation_Marker = "Fail"
            End If
            System.Threading.Thread.Sleep(500)
            frmMANUALTEST.Refresh()
        Else
            If SwitchPorts = 1 Then ' ~~~~~~~ 6 port stuff ~~~~~~~~~
                If SwitchedChecked Then  'Auto RF Switching
                    If SetSwitchesPosition <> 2 Then
                        SetSwitchesPosition = 2
                        status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)
                    End If
                    frmMANUALTEST.Switch(1)
                Else
                    MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
                End If
            Else ' ~~~~~~~ 4 port stuff ~~~~~~~~~
                If SwitchedChecked Then  'Auto RF Switching
                    If SetSwitchesPosition <> 3 Then
                        SetSwitchesPosition = 3
                        status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)
                    End If
                    frmMANUALTEST.Switch(2)
                Else
                    MsgBox("Move Cables to RF Position 3", vbOKOnly, "Manual Switch")
                End If
            End If

            If MutiCalChecked Then SetupVNA(True, 3)

            System.Threading.Thread.Sleep(100)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & 0 - GetSpecification("Isolation"))
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 2")
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                ExtraAvg()
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                ISo = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":CALC:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & 0 - GetSpecification("Isolation"))
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 2")
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                ExtraAvg()
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                ISo = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
            Else
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("OPC?;LOGM;")
                ScanGPIB.BusWrite("OPC?;REFV " & 0 - GetSpecification("Isolation"))
                ScanGPIB.BusWrite("OPC?;SCAL 2")
                ScanGPIB.BusWrite("MARK1;")  'Marker 1 on
                ExtraAvg(2)
                ScanGPIB.BusWrite("MARKMAXI;")  'Marker 1 max
                ISo = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
            End If
            ISo = ISo + CDbl(frmMANUALTEST.txtOffset3.Text)
            frmMANUALTEST.Refresh()
            Spec = 0 - GetSpecification("Isolation", freq)
            'ISo = Format(ISo, 0.00)
            If ISo <= Spec Then
                Isolation_Marker = "Pass"
            Else
                Isolation_Marker = "Fail"
            End If
        End If
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
        Else
            ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
        End If
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function
    Public Function IsolationCOMB(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim MaxData(32) As Double
        Dim MaxData1(32) As Double
        Dim MaxData2(32) As Double
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace1Freq(Points) As Double
        Dim LowerTraceData(Points) As Double
        Dim UpperTraceData(Points) As Double
        Dim returnValue1 As Double
        Dim returnValue2 As Double
        Dim TraceID1 As Long
        Dim x, y, z, A, b As Long
        Dim t1 As Trace
        Dim T1Min As Double
        Dim T1Max As Double
        Dim Workstation As String
        Dim Title As String
        Dim NumPorts As Double
        Dim PortNum As Byte
        Dim Ports As Integer

        Title = ActiveTitle
        ActiveTitle = "     TESTING ISOLATION       "
        IsolationCOMB = ""
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset3.Text = "" Then frmMANUALTEST.txtOffset3.Text = 0 ' bad user protection
        t1 = New Trace
        Spec = 0 - GetSpecification("Isolation")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset3.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Return "Pass"
            Else
                Return "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                TraceID1 = 4265
                GetTracePoints(TraceID1)
                T1Min = MinNoZero(YArray)
                T1Max = MaxNoZero(YArray)
                ISo = T1Max
                ISo = ISo + CDbl(frmMANUALTEST.txtOffset3.Text)

                If ISo < Spec Then
                    Return "Pass"
                Else
                    Return "Fail"
                End If
            ElseIf PassChecked Then
                ISo = SpecISO
                Return "Pass"
            ElseIf FailChecked Then
                ISo = SpecISO + 10
                Return "Fail"
            End If
            System.Threading.Thread.Sleep(500)
            frmMANUALTEST.Refresh()
        Else
            ILSetDone = False
            MsgBox("Turn the Combiner/Divider around. Connect J2 To VNA Port1. Add 50 OHM load to J1", , "Isolation Test")
            NumPorts = GetSpecification("Ports")
            PortNum = CByte(NumPorts)
            Ports = Int(NumPorts)
            x = 1

            For b = 2 To Ports
                PortNum = CByte(b)
                ActiveTitle = "     TESTING Isolation    SW POSITION " & x & "      "
                If SwitchedChecked Then  'Auto RF Switching
                    SetSwitchesPosition = PortNum
                    status = SwitchCom.SetSwitchPosition(PortNum) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                    frmMANUALTEST.Switch(b - 1)
                Else
                    MsgBox("Move Cables to RF Position " & b & " ")
                End If

                If MutiCalChecked Then SetupVNA(True, x)
                frmMANUALTEST.Refresh()
                If Not ILSetDone Then
                    If VNAStr = "AG_E5071B" Then
                        ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                        ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                    ElseIf VNAStr = "N3383A" Then
                        ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                        ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                    Else
                        ScanGPIB.BusWrite("OPC?;CHAN2;")
                        ScanGPIB.BusWrite("OPC?;LOGM;")
                        ScanGPIB.BusWrite("OPC?;REFV 0")
                        ScanGPIB.BusWrite("OPC?;SCAL 10")
                    End If
                    ILSetDone = True
                End If
                frmMANUALTEST.Refresh()
                ExtraAvg()
                If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                    Title = "Manual " & TestRun & "Isolation Port " & Ports & " "
                    SerialNumber = "UUT" & UUTNum_Reset
                    TestID = TestID
                    CalDate = Now
                    Notes = ""
                    Workstation = GetComputerName()
                    TraceID1 = SQL.GetTraceID(Title, TestID)
                    TraceID = TraceID1
                End If
                ScanGPIB.GetTrace(Trace1Freq, IL1Data)
                Trace1Freq = TrimX(Trace1Freq)
                IL1Data_offs = TrimY(IL1Data, CDbl(frmMANUALTEST.txtOffset3.Text))
                If TraceChecked And Not TweakMode Then
                    If UUTNum_Reset <= 5 Then
                        ReDim Preserve XArray(TraceData.Count - 1)
                        ReDim Preserve YArray(TraceData.Count - 1)
                        Array.Clear(YArray, 0, TraceData.Count - 1)
                        XArray = TraceFreq
                        YArray = IL1Data
                        SQL.SaveTrace(Title, TestID, TraceID)
                        YArray = IL1Data_offs
                    End If
                End If
                SpecCuttoffFreq = 0
                SpecCuttoffFreq = SQL.GetSpecification("CutOffFreqMHz")
                If ISO_TF Then
                    For z = 0 To UBound(Trace1Freq)
                        If Trace1Freq(z) >= (SpecCuttoffFreq) Then GoTo SetPoints
                    Next z
SetPoints:
                    frmMANUALTEST.Refresh()
                    For y = 0 To z - 1
                        ReDim Preserve LowerTraceData(y)
                        LowerTraceData(y) = IL1Data(y)
                    Next y
                    ReDim Preserve MaxData1(b - 2)
                    MaxData1(b - 2) = MaxNoZero(LowerTraceData)

                    z = 0
                    For A = (y + 1) To UBound(Trace1Freq)
                        ReDim Preserve UpperTraceData(z)
                        UpperTraceData(z) = IL1Data(A)
                        z = z + 1
                    Next A
                    ReDim Preserve MaxData2(b - 2)
                    MaxData2(b - 2) = MaxNoZero(UpperTraceData)
                    If b = 2 Then
                        returnValue1 = MaxData1(b - 2)
                        returnValue2 = MaxData2(b - 2)
                    Else
                        If MaxData1(b - 2) > returnValue1 Then returnValue1 = MaxData1(b - 2)
                        If MaxData2(b - 2) > returnValue2 Then returnValue2 = MaxData2(b - 2)
                    End If

                    ISoL = returnValue1
                    ISoH = returnValue2

                Else
                    ReDim Preserve MaxData(b - 2)
                    MaxData(b - 2) = MaxNoZero(IL1Data)
                    If b = 2 Then
                        ISo = MaxData(b - 2)
                    Else
                        If MaxData(b - 2) > ISo Then ISo = MaxData(b - 2)
                    End If

                End If
                Pts = Points
                x = x + 1
            Next b
            If ISO_TF Then
                frmMANUALTEST.Spec3Min.Text = "N/A"
                ISoL = ISoL + CDbl(frmMANUALTEST.txtOffset3.Text)
                ISoL = TruncateDecimal(ISoL, 1)
                ISoH = ISoH + CDbl(frmMANUALTEST.txtOffset3.Text)
                ISoH = TruncateDecimal(ISoH, 1)
                If ISoL <= SpecISOL And ISoH <= SpecISOH Then
                    IsolationCOMB = "Pass"
                Else
                    IsolationCOMB = "Fail"
                End If
            Else
                frmMANUALTEST.Spec3Min.Text = "N/A"
                ISo = ISo + CDbl(frmMANUALTEST.txtOffset3.Text)
                ISo = TruncateDecimal(ISo, 1)
                If ISo <= SpecISO Then
                    IsolationCOMB = "Pass"
                Else
                    IsolationCOMB = "Fail"
                End If
            End If
        End If
        frmMANUALTEST.Refresh()
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function

    Public Function IsolationCOMB_Marker(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim MaxData(32) As Double
        Dim MaxData1(32) As Double
        Dim MaxData2(32) As Double
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace1Freq(Points) As Double
        Dim LowerTraceData(Points) As Double
        Dim UpperTraceData(Points) As Double
        Dim TraceID1 As Long
        Dim x, b As Long
        Dim t1 As Trace
        Dim T1Min As Double
        Dim T1Max As Double
        Dim Workstation As String
        Dim Title As String
        Dim NumPorts As Double
        Dim PortNum As Byte
        Dim Ports As Integer

        Title = ActiveTitle
        ActiveTitle = "     TESTING ISOLATION       "
        IsolationCOMB_Marker = ""
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset3.Text = "" Then frmMANUALTEST.txtOffset3.Text = 0 ' bad user protection
        t1 = New Trace
        Spec = 0 - GetSpecification("Isolation")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset3.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Return "Pass"
            Else
                Return "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                TraceID1 = 4265
                GetTracePoints(TraceID1)
                T1Min = MinNoZero(YArray)
                T1Max = MaxNoZero(YArray)
                ISo = T1Max
                ISo = ISo + CDbl(frmMANUALTEST.txtOffset3.Text)

                If ISo < Spec Then
                    Return "Pass"
                Else
                    Return "Fail"
                End If
            ElseIf PassChecked Then
                ISo = SpecISO
                Return "Pass"
            ElseIf FailChecked Then
                ISo = SpecISO + 10
                Return "Fail"
            End If
            System.Threading.Thread.Sleep(500)
            frmMANUALTEST.Refresh()
        Else
            ILSetDone = False
            MsgBox("Turn the Combiner/Divider around. Connect J2 To VNA Port1. Add 50 OHM load to J1", , "Isolation Test")
            NumPorts = GetSpecification("Ports")
            PortNum = CByte(NumPorts)
            Ports = Int(NumPorts)
            x = 0

            For b = 2 To Ports
                PortNum = CByte(b)
                ActiveTitle = "     TESTING Isolation    SW POSITION " & x & "      "
                If SwitchedChecked Then  'Auto RF Switching
                    SetSwitchesPosition = PortNum
                    status = SwitchCom.SetSwitchPosition(PortNum) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                    frmMANUALTEST.Switch(b - 1)
                Else
                    MsgBox("Move Cables to RF Position " & b & " ")
                End If

                If MutiCalChecked Then SetupVNA(True, b)
                frmMANUALTEST.Refresh()
                If Not ILSetDone Then
                    If VNAStr = "AG_E5071B" Then
                        ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                        ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                        ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                        ExtraAvg()
                        ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                        ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                        ISo = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                    ElseIf VNAStr = "N3383A" Then
                        ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                        ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                        ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                        ExtraAvg()
                        ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                        ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                        ISo = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                    Else
                        ScanGPIB.BusWrite("OPC?;CHAN2;")
                        ScanGPIB.BusWrite("OPC?;LOGM;")
                        ScanGPIB.BusWrite("OPC?;REFV 0")
                        ScanGPIB.BusWrite("OPC?;SCAL 10")
                        ScanGPIB.BusWrite("MARK1;")  'Marker 1 on
                        Delay(500)
                        ExtraAvg(2)
                        ScanGPIB.BusWrite("MARKMAXI;")  'Marker 1 max
                        ISo = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
                    End If
                    ILSetDone = True
                End If
                frmMANUALTEST.Refresh()
                SpecCuttoffFreq = 0
                SpecCuttoffFreq = SQL.GetSpecification("CutOffFreqMHz")
                ReDim Preserve MaxData(x)
                MaxData(x) = ISo
                If x > 0 Then
                    If MaxData(x - 1) < ISo Then ISo = MaxData(x - 1)
                End If
                Pts = Points
                x = x + 1
            Next b
            ISo = ISo + CDbl(frmMANUALTEST.txtOffset3.Text)
            ISo = TruncateDecimal(ISo, 1)
            If ISo <= Spec Then
                IsolationCOMB_Marker = "Pass"
            Else
                IsolationCOMB_Marker = "Fail"
            End If
        End If
        frmMANUALTEST.Refresh()
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
        'Turn off all markers
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
            ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
            ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
        Else
            ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
        End If
    End Function

    Public Function Coupling(TestRun As String, Direction As Long, SpecType As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim SpecPM As Double
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim TraceID1 As Long
        Dim x As Long
        Dim t1 As Trace
        Dim T1Min As Double
        Dim T1Max As Double
        Dim Workstation As String
        Dim Title As String
        Dim Direction1Value As Double
        Dim Direction2Value As Double
        Dim J3, J4 As Double

        Coupling = ""
        Title = ActiveTitle
        If Direction = 1 Then ActiveTitle = "     TESTING COUPLING   FORWARD DIRECTION    "
        If Direction = 2 Then ActiveTitle = "     TESTING COUPLING  REVERSE DIRECTION    "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset3.Text = "" Then frmMANUALTEST.txtOffset3.Text = 0 ' bad user protection
        t1 = New Trace
        Spec = GetSpecification("Coupling")
        SpecPM = GetSpecification("CouplingPM")
        SwitchPorts = SQL.GetSpecification("SwitchPorts")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset2.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Coupling = "Pass"
            Else
                Coupling = "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                TraceID1 = 4265
                GetTracePoints(TraceID1)
                T1Min = MinNoZero(YArray)
                T1Max = MaxNoZero(YArray)
                COuP = T1Max
                COuP = COuP + CDbl(frmMANUALTEST.txtOffset3.Text)

                If COuP < Spec Then
                    Coupling = "Pass"
                Else
                    Coupling = "Fail"
                End If
            ElseIf PassChecked Then
                COuP = Spec
                Coupling = "Pass"
            ElseIf FailChecked Then
                COuP = Spec + 10
                Coupling = "Fail"
            End If
            System.Threading.Thread.Sleep(500)
            frmMANUALTEST.Refresh()
        Else

            If SwitchPorts Then
                If Direction = 1 Then
                    If SwitchedChecked Then  'Auto RF Switching
                        If SetSwitchesPosition <> 2 Then
                            SetSwitchesPosition = 2
                            status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                            System.Threading.Thread.Sleep(500)
                        End If
                        frmMANUALTEST.Switch(1)
                    Else
                        MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
                    End If
                    If MutiCalChecked Then SetupVNA(True, 1)
                ElseIf Direction = 2 Then
                    If SwitchedChecked Then  'Auto RF Switching
                        If SetSwitchesPosition <> 4 Then
                            SetSwitchesPosition = 4
                            status = SwitchCom.SetSwitchPosition(4) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                            System.Threading.Thread.Sleep(500)
                        End If
                        frmMANUALTEST.Switch(3)
                    Else
                        MsgBox("Move Cables to RF Position 4", vbOKOnly, "Manual Switch")
                    End If
                    If MutiCalChecked Then SetupVNA(True, 1)
                End If
            Else
                If Direction = 1 Then
                    If SwitchedChecked Then  'Auto RF Switching
                        If SetSwitchesPosition <> 2 Then
                            SetSwitchesPosition = 2
                            status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                            System.Threading.Thread.Sleep(500)
                        End If
                        frmMANUALTEST.Switch(1)
                    Else
                        MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
                    End If
                    If MutiCalChecked Then SetupVNA(True, 1)
                ElseIf Direction = 2 Then
                    If SwitchedChecked Then  'Auto RF Switching
                        If SetSwitchesPosition <> 2 Then
                            SetSwitchesPosition = 2
                            status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                            System.Threading.Thread.Sleep(500)
                        End If
                        frmMANUALTEST.Switch(1)
                    Else
                        MsgBox("Move Cables to RF Position 1", vbOKOnly, "Manual Switch")
                    End If
                    If MutiCalChecked Then SetupVNA(True, 1)
                End If
            End If


            frmMANUALTEST.Refresh()

            If SpecType = "DUAL DIRECTIONAL COUPLER" Or ((SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER") And Direction = 1) Then
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & 0 - GetSpecification("Coupling"))
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 0.5")
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & 0 - GetSpecification("Coupling"))
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 0.5")
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                    ScanGPIB.BusWrite("OPC?;LOGM;")
                    'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                    ScanGPIB.BusWrite("OPC?;REFV " & 0 - GetSpecification("Coupling"))
                    'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                    ScanGPIB.BusWrite("OPC?;SCAL 0.5")
                    'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                End If

                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or (SpecType = "DUAL DIRECTIONAL COUPLER" And Direction = 1) Then Title = "Manual " & TestRun & "Coupling J3"
                If SpecType = "DUAL DIRECTIONAL COUPLER" And Direction = 2 Then Title = "Manual " & TestRun & "Coupling J4"

                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or (SpecType = "DUAL DIRECTIONAL COUPLER" And Direction = 1) Then
                    If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                        SerialNumber = "UUT" & UUTNum_Reset
                        TestID = TestID
                        CalDate = Now
                        Notes = ""
                        Workstation = GetComputerName()
                        TraceID1 = SQL.GetTraceID(Title, TestID)
                        TraceID = TraceID1
                    End If
                    ScanGPIB.GetTrace(TraceFreq, COUP1Data)
                    TraceFreq = TrimX(TraceFreq)
                    COUP1Data_offs = TrimY(COUP1Data, CDbl(frmMANUALTEST.txtOffset3.Text))
                    If TraceChecked And Not TweakMode Then
                        If UUTNum_Reset <= 5 Then
                            ReDim Preserve XArray(TraceFreq.Count - 1)
                            ReDim Preserve YArray(TraceFreq.Count - 1)
                            Array.Clear(YArray, 0, TraceFreq.Count - 1)
                            XArray = TraceFreq
                            YArray = COUP1Data
                            SQL.SaveTrace(Title, TestID, TraceID)
                            YArray = COUP1Data_offs
                            Try
                                For x = 0 To YArray.Count - 1
                                    COUP_XArray(UUTNum_Reset - 1, x) = XArray(x)
                                    COUP1_YArray(UUTNum_Reset - 1, x) = YArray(x)
                                Next
                            Catch ex As Exception

                            End Try
                        End If
                    End If
                Else
                    If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                        SerialNumber = "UUT" & UUTNum_Reset
                        TestID = TestID
                        CalDate = Now
                        Notes = ""
                        Workstation = GetComputerName()
                        TraceID1 = SQL.GetTraceID(Title, TestID)
                        TraceID = TraceID1
                    End If
                    ScanGPIB.GetTrace(TraceFreq, COUP2Data)
                    TraceFreq = TrimX(TraceFreq)
                    COUP2Data_offs = TrimY(COUP2Data, CDbl(frmMANUALTEST.txtOffset3.Text))

                    If TraceChecked And Not TweakMode Then
                        ReDim Preserve XArray(TraceFreq.Count - 1)
                        ReDim Preserve YArray(TraceFreq.Count - 1)
                        Array.Clear(YArray, 0, TraceFreq.Count - 1)
                        YArray = COUP2Data
                        SQL.SaveTrace(Title, TestID, TraceID)
                        YArray = COUP2Data_offs
                        If UUTNum_Reset <= 5 Then
                            For x = 0 To YArray.Count - 1
                                COUP2_YArray(UUTNum_Reset - 1, x) = YArray(x)
                            Next
                        End If
                    End If
                End If
            End If

            frmMANUALTEST.Refresh()
            If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or Direction = 1 Then
                For x = 0 To 200
                    If x = 0 Then COuP = Math.Abs(COUP1Data(x))
                    If x <> 0 Then COuP = COuP + Math.Abs(COUP1Data(x))
                Next x
                COuP = COuP / 201

                COuP = COuP + CDbl(frmMANUALTEST.txtOffset3.Text)

                If COuP <= Spec + SpecPM And COuP >= Spec - SpecPM Then
                    Coupling = "Pass"
                Else
                    Coupling = "Fail"
                End If
            Else
                For x = 0 To 200
                    If x = 0 Then Direction1Value = Math.Abs(COUP1Data(x))
                    If x <> 0 Then Direction1Value = Direction1Value + Math.Abs(COUP1Data(x))
                Next x
                Direction1Value = Direction1Value / 201
                J3 = Math.Abs(Spec - Direction1Value)

                For x = 0 To 200
                    If x = 0 Then Direction2Value = Math.Abs(COUP2Data(x))
                    If x <> 0 Then Direction2Value = Direction2Value + Math.Abs(COUP2Data(x))
                Next x
                Direction2Value = Direction2Value / 201
                J4 = Math.Abs(Spec - Direction2Value)

                If J3 > J4 Then
                    COuP = Direction1Value
                Else
                    COuP = Direction2Value
                End If

                COuP = COuP + CDbl(frmMANUALTEST.txtOffset3.Text)

                If COuP <= (Spec + SpecPM) And COuP >= (Spec - SpecPM) Then
                    Coupling = "Pass"
                Else
                    Coupling = "Fail"
                End If
            End If
        End If
        COuP = TruncateDecimal(COuP, 1) + frmMANUALTEST.txtOffset3.Text
        COuP = TruncateDecimal(COuP, 1)
        ActiveTitle = Title
        frmMANUALTEST.Refresh()
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function


    Public Function AmplitudeBalance(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 0) As String

        Dim Spec As Double
        Dim TraceData(Points) As Object
        Dim TraceFreq(Points) As Object
        Dim Trace1Data(Points) As Double
        Dim Trace1Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim t1 As Trace
        Dim ABMin As Double
        Dim ABMax As Double
        Dim ABArray(Points) As Double
        Dim ABArray1(Points) As Double
        Dim ABArray2(Points) As Double
        Dim i As Integer
        Dim ABStr As String
        Dim ABMin2 As Double
        Dim ABStrArray
        Dim Workstation As String
        Dim Title As String

        Title = ActiveTitle
        ActiveTitle = "     TESTING AMPLITUDE BALANCE     "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset4.Text = "" Then frmMANUALTEST.txtOffset4.Text = 0
        t1 = New Trace
        Spec = GetSpecification("AmplitudeBalance")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset4.Text)
            AB = RetrnVal
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                AmplitudeBalance = "Pass"
            Else
                AmplitudeBalance = "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                ABTraceID1 = 9889
                ABTraceID2 = 9890
                GetTracePoints(ABTraceID1)
                Trace1Data = YArray
                GetTracePoints(ABTraceID2)
                Trace2Data = YArray


                For i = 0 To Pts
                    ABArray1(i) = Trace1Data(i)
                    ABArray2(i) = Trace2Data(i)
                Next

                For i = 0 To Pts
                    ABArray(i) = Math.Abs(ABArray1(i)) - Math.Abs(ABArray2(i)) / 2
                    ABMax = (ABArray1(i) - ABArray2(i)) / 2
                Next

                frmMANUALTEST.Refresh()

                ABMax = MaxNoZero(ABArray)
                i = MaxX(ABArray)

                ABMin2 = ABArray2(i)

                AB = (ABMax - ABMin) / 2
                If AB < Spec Then
                    AmplitudeBalance = "Pass"
                Else
                    AmplitudeBalance = "Fail"
                End If

                ABStr = Str(AB)
                ABStrArray = Split(AB, ".")

                'AB = CDbl((ABStrArray(0) & "." & Left(ABStrArray(1), 3)))
                'AB = AB + CDbl(frmMANUALTEST.txtOffset4.Text)
                AB = TruncateDecimal(AB, 2)
            ElseIf PassChecked Then
                AB = Spec
                AmplitudeBalance = "Pass"
            ElseIf FailChecked Then
                AB = Spec + 10
                AmplitudeBalance = "Fail"
            End If
        Else
            frmMANUALTEST.Refresh()
            For i = 0 To Pts - 1
                ABArray(i) = Math.Abs(Math.Abs(IL1Data(i)) - Math.Abs(IL2Data(i)))
            Next

            AB = MaxNoZero(ABArray) / 2
            If AB < Spec Then
                AmplitudeBalance = "Pass"
            Else
                AmplitudeBalance = "Fail"
            End If

            ABStr = CStr(AB)
            ABStrArray = Split(AB, ".")
            If ABStrArray(0) = "0" Then GoTo Round
            AB = CDbl((ABStrArray(0) & "." & Left(ABStrArray(1), 3)))
        End If

        frmMANUALTEST.Refresh()
Round:
        AB = TruncateDecimal(AB, 2)
        AB = Format(Math.Round(AB + CDbl(frmMANUALTEST.txtOffset4.Text), 2), "0.00")
        If AB <= Spec Then
            AmplitudeBalance = "Pass"
        Else
            AB = Spec + 10
            AmplitudeBalance = "Fail"
        End If

        ActiveTitle = Title
    End Function

    Public Function AmplitudeBalance_Marker(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 0) As String

        Dim Spec As Double
        Dim TraceData(Points) As Object
        Dim TraceFreq(Points) As Object
        Dim Trace1Data(Points) As Double
        Dim Trace1Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim t1 As Trace
        Dim ABMin As Double
        Dim ABMax As Double
        Dim ABArray(Points) As Double
        Dim ABArray1(Points) As Double
        Dim ABArray2(Points) As Double
        Dim i As Integer
        Dim ABStr As String
        Dim ABMin2 As Double
        Dim ABStrArray
        Dim Workstation As String
        Dim Title As String

        Title = ActiveTitle
        ActiveTitle = "     TESTING AMPLITUDE BALANCE     "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset4.Text = "" Then frmMANUALTEST.txtOffset4.Text = 0
        t1 = New Trace
        Spec = GetSpecification("AmplitudeBalance")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset4.Text)
            AB = RetrnVal
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                AmplitudeBalance_Marker = "Pass"
            Else
                AmplitudeBalance_Marker = "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                ABTraceID1 = 9889
                ABTraceID2 = 9890
                GetTracePoints(ABTraceID1)
                Trace1Data = YArray
                GetTracePoints(ABTraceID2)
                Trace2Data = YArray

                For i = 0 To Pts
                    ABArray1(i) = Trace1Data(i)
                    ABArray2(i) = Trace2Data(i)
                Next

                For i = 0 To Pts
                    ABArray(i) = Math.Abs(ABArray1(i)) - Math.Abs(ABArray2(i)) / 2
                    ABMax = (ABArray1(i) - ABArray2(i)) / 2
                Next

                frmMANUALTEST.Refresh()

                ABMax = MaxNoZero(ABArray)
                i = MaxX(ABArray)

                ABMin2 = ABArray2(i)

                AB = (ABMax - ABMin) / 2
                If AB < Spec Then
                    AmplitudeBalance_Marker = "Pass"
                Else
                    AmplitudeBalance_Marker = "Fail"
                End If

                ABStr = Str(AB)
                ABStrArray = Split(AB, ".")

                'AB = CDbl((ABStrArray(0) & "." & Left(ABStrArray(1), 3)))
                'AB = AB + CDbl(frmMANUALTEST.txtOffset4.Text)
                AB = TruncateDecimal(AB, 2)
            ElseIf PassChecked Then
                AB = Spec
                AmplitudeBalance_Marker = "Pass"
                AB = Spec + 10
                AmplitudeBalance_Marker = "Fail"
            ElseIf FailChecked Then
                AB = Spec + 10
                AmplitudeBalance_Marker = "Fail"
            End If
        Else
            frmMANUALTEST.Refresh()

            If Math.Abs(IL1AB) > Math.Abs(IL2AB) Then
                AB = Math.Abs(IL1AB)
            Else
                AB = Math.Abs(IL2AB)
            End If
            AB = AB / 2

            If AB < Spec Then
                AmplitudeBalance_Marker = "Pass"
            Else
                AmplitudeBalance_Marker = "Fail"
            End If

        End If

        frmMANUALTEST.Refresh()
Round:
        AB = TruncateDecimal(AB, 3)
        AB = Format(Math.Round(AB + CDbl(frmMANUALTEST.txtOffset4.Text), 3), "0.00")
        If AB <= Spec Then
            AmplitudeBalance_Marker = "Pass"
        Else
            AmplitudeBalance_Marker = "Fail"
        End If
        'Turn off all markers
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
            ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
            ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
        Else
            ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
        End If
        ActiveTitle = Title
    End Function

    Public Function AmplitudeBalance_multiband(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 0) As String
        Dim Spec As Double
        Dim TraceData(Points) As Object
        Dim TraceFreq(Points) As Object
        Dim Trace1Data(Points) As Double
        Dim Trace1Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim t1 As Trace
        Dim ABMin As Double
        Dim ABMax As Double
        Dim ABArray(Points) As Double
        Dim ABArray1(Points) As Double
        Dim ABArray2(Points) As Double
        Dim ABArray_ext(Points) As Double
        Dim ABArray1_ext(Points) As Double
        Dim ABArray2_ext(Points) As Double
        Dim i As Integer
        Dim ABStr As String
        Dim ABMin2 As Double
        Dim ABStrArray
        Dim Workstation As String
        Dim Title As String
        AmplitudeBalance_multiband = "Fail"
        Title = ActiveTitle
        ActiveTitle = "     TESTING AMPLITUDE BALANCE     "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset4.Text = "" Then frmMANUALTEST.txtOffset4.Text = 0
        t1 = New Trace
        Spec = GetSpecification("AmplitudeBalance")
        If ResumeTesting Then
            AB1 = AB1 + CDbl(frmMANUALTEST.txtOffset4.Text)
            AB2 = AB2 + CDbl(frmMANUALTEST.txtOffset4.Text)
            If AB1 > AB2 Then
                RetrnVal = AB1
            Else
                RetrnVal = AB2
            End If
            AB = RetrnVal
            If AB1 <= Spec Then
                AB1Pass = "Pass"
            Else
                AB1Pass = "Fail"
            End If
            If AB2 <= SpecAB_exp Then
                AB2Pass = "Pass"
            Else
                AB2Pass = "Fail"
            End If
            If AB1Pass = "Pass" And AB1Pass = "Pass" Then
                AmplitudeBalance_multiband = "Pass"
            Else
                AmplitudeBalance_multiband = "Fail"
            End If

        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                ABTraceID1 = 9889
                ABTraceID2 = 9890
                GetTracePoints(ABTraceID1)
                Trace1Data = YArray
                GetTracePoints(ABTraceID2)
                Trace2Data = YArray

                For i = 0 To Pts
                    ABArray1(i) = Trace1Data(i)
                    ABArray2(i) = Trace2Data(i)
                Next

                For i = 0 To Pts
                    ABArray(i) = Math.Abs(ABArray1(i)) - Math.Abs(ABArray2(i)) / 2
                    ABMax = (ABArray1(i) - ABArray2(i)) / 2
                Next

                frmMANUALTEST.Refresh()

                ABMax = MaxNoZero(ABArray)
                i = MaxX(ABArray)

                ABMin2 = ABArray2(i)

                AB = (ABMax - ABMin) / 2
                If AB < Spec Then
                    AmplitudeBalance_multiband = "Pass"
                Else
                    AmplitudeBalance_multiband = "Fail"
                End If

                ABStr = Str(AB)
                ABStrArray = Split(AB, ".")

                'AB = CDbl((ABStrArray(0) & "." & Left(ABStrArray(1), 3)))
                'AB = AB + CDbl(frmMANUALTEST.txtOffset4.Text)
                AB = TruncateDecimal(AB, 2)
            ElseIf PassChecked Then
                If SpecAB_TF Then
                    AB1 = Spec
                    AB2 = Spec
                    AB1Pass = "Pass"
                    AB2Pass = "Pass"
                Else
                    AB = Spec
                End If
                AmplitudeBalance_multiband = "Pass"
            ElseIf FailChecked Then
                If SpecAB_TF Then
                    AB1 = Spec + 10
                    AB2 = Spec + 10
                    AB1Pass = "Fail"
                    AB2Pass = "Fail"
                Else
                    AB = Spec + 10
                End If
                AmplitudeBalance_multiband = "Fail"
            End If
        Else
            frmMANUALTEST.Refresh()
            For i = 0 To IL1Data_offs1.Count - 2
                ReDim Preserve ABArray1(i)
                ABArray1(i) = Math.Abs(Math.Abs(IL1Data_offs1(i)) - Math.Abs(IL2Data_offs1(i)))
            Next
            AB1 = MaxNoZero(ABArray1) / 2
            Dim test_ab1 = AB1 + CDbl(frmMANUALTEST.txtOffset4.Text)
            If test_ab1 <= Spec Then
                AmplitudeBalance_multiband = "Pass"
                AB1Pass = "Pass"
            Else
                AmplitudeBalance_multiband = "Fail"
                AB1Pass = "Fail"
            End If
            For i = 0 To IL2Data_offs2.Count - 2
                ReDim Preserve ABArray1(i)
                ABArray2(i) = Math.Abs(Math.Abs(IL1Data_offs2(i)) - Math.Abs(IL2Data_offs2(i)))
            Next
            AB2 = MaxNoZero(ABArray2) / 2
            Dim test_ab2 = AB2 + CDbl(frmMANUALTEST.txtOffset4.Text)
            If test_ab2 <= SpecAB_exp Then
                AB2Pass = "Pass"
            Else
                AB2Pass = "Fail"
                AmplitudeBalance_multiband = "Fail"
            End If
            If test_ab2 <= SpecAB_exp And AmplitudeBalance_multiband = "Pass" Then
                AmplitudeBalance_multiband = "Pass"
            Else
                AmplitudeBalance_multiband = "Fail"
            End If

            If AB1 > AB2 Then
                ABStr = CStr(AB1)
                ABStrArray = Split(ABStr, ".")
                If ABStrArray(0) = "0" Then GoTo Round
                AB = CDbl((ABStrArray(0) & "." & Left(ABStrArray(1), 3)))
            Else
                ABStr = CStr(AB2)
                ABStrArray = Split(ABStr, ".")
                If ABStrArray(0) = "0" Then GoTo Round
                AB = CDbl((ABStrArray(0) & "." & Left(ABStrArray(1), 3)))
            End If
        End If
        frmMANUALTEST.Refresh()
Round:
        AB = TruncateDecimal(AB, 2)
        AB = Format(Math.Round(AB + CDbl(frmMANUALTEST.txtOffset4.Text), 3), "0.00")
        If SpecAB_TF And Not ResumeTesting Then
            If AB1 > AB2 Then
                AB = AB1
            Else
                AB = AB2
            End If
            AB1 = TruncateDecimal(AB1, 2)
            AB1 = Format(TruncateDecimal(AB1 + CDbl(frmMANUALTEST.txtOffset4.Text), 3), "0.00")
            AB2 = TruncateDecimal(AB2, 2)
            AB2 = Format(TruncateDecimal(AB2 + CDbl(frmMANUALTEST.txtOffset4.Text), 3), "0.00")
        End If


        ActiveTitle = Title
        SetSwitchesPosition = 1
        Status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function

    Public Function AmplitudeBalanceCOMB(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 0) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceData(Points) As Object
        Dim TraceFreq(Points) As Object
        Dim MaxData(32) As Double
        Dim MinData(32) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace1Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim TraceID1 As Long
        Dim x As Long
        Dim t1 As Trace
        Dim ABMin As Double
        Dim ABMax As Double
        Dim ABArray(Points) As Double
        Dim i As Integer
        Dim Workstation As String
        Dim Title As String
        Dim NumPorts As Double
        Dim PortNum As Byte
        Dim Ports As Integer
        Dim ABStr As String = ""
        Dim ABStrArray(500) As String


        Title = ActiveTitle
        AmplitudeBalanceCOMB = ""
        ABSetDone = False
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset4.Text = "" Then frmMANUALTEST.txtOffset4.Text = 0
        t1 = New Trace
        Spec = GetSpecification("AmplitudeBalance")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset4.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Return "Pass"
            Else
                Return "Fail"
            End If
            frmMANUALTEST.Refresh()
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                ABTraceID1 = 4262
                ABTraceID2 = 4263
                GetTracePoints(ABTraceID1)

                ABArray = YArray

                ABMax = MaxNoZero(ABArray)
                i = MaxX(ABArray)

                GetTracePoints(ABTraceID2)
                ABArray = YArray
                ABMin = Min(YArray)

                AB = (ABMax - ABMin) / 2
                If AB < Spec Then
                    Return "Pass"
                Else
                    Return "Fail"
                End If

                ABStr = Str(AB)
                ABStrArray = Split(AB, ".")

                AB = CDbl((ABStrArray(0) & "." & Left(ABStrArray(1), 3)))
                AB = AB + CDbl(frmMANUALTEST.txtOffset4.Text)
                AB = TruncateDecimal(AB, 2)
            ElseIf PassChecked Then
                AB = Spec
                Return "Pass"
            ElseIf FailChecked Then
                AB = Spec + 10
                Return "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            NumPorts = GetSpecification("Ports")
            PortNum = CByte(NumPorts)
            Ports = Int(NumPorts)

            For x = 1 To Ports
                PortNum = CByte(x)
                ActiveTitle = "     TESTING AMPLITUDE BALANCE     SW POSITION " & x & "      "
                If SwitchedChecked Then  'Auto RF Switching
                    SetSwitchesPosition = PortNum
                    status = SwitchCom.SetSwitchPosition(PortNum) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                    frmMANUALTEST.Switch(x - 1)
                Else
                    MsgBox("Move Cables to RF Position " & x & " ")
                End If

                If MutiCalChecked Then
                    SetupVNA(True, 2)
                    If x <> 1 Then
                        If VNAStr = "AG_E5071B" Then
                            ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                            ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                            ScanGPIB.BusWrite(":INIT2:CONT ON") ' and start another sweep
                        ElseIf VNAStr = "N3383A" Then
                            ScanGPIB.BusWrite("CALC:PAR:SEL 'CH1_S21_1'")
                            ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                            ScanGPIB.BusWrite(":INIT1:CONT ON") ' and start another sweep
                        Else
                            ScanGPIB.BusWrite("OPC?;INPUDATA;") ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(gBuffer & ";")
                            ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
                            'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                            ScanGPIB.BusWrite("OPC?;CONT") ' and start another sweep
                            'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                        End If
                    End If
                End If
                If Not ABSetDone Or MutiCalChecked Then
                    If VNAStr = "AG_E5071B" Then
                        ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                        ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                    ElseIf VNAStr = "N3383A" Then
                        ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory Off"
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON")  ' Data On
                        ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                    Else
                        ScanGPIB.BusWrite("OPC?;DISPDATA;")
                        'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                        ScanGPIB.BusWrite("OPC?;CHAN2;")
                        'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                        ScanGPIB.BusWrite("OPC?;LOGM;")
                        'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                        ScanGPIB.BusWrite("OPC?;REFV 0")
                        'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                        ScanGPIB.BusWrite("OPC?;SCAL 10")
                        'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                    End If
                    ABSetDone = True
                End If
                If x = 1 Then If VNAStr = "AG_E5071B" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                If x = 1 Then If VNAStr = "N3383A" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                If x = 1 Then If VNAStr <> "AG_E5071B" And VNAStr <> "N3383A" Then ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory

                frmMANUALTEST.Refresh()
                If x <> 1 Then
                    If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                        Title = "Manual " & TestRun & "AmplitudeBalance Port " & PortNum
                        SerialNumber = "UUT" & UUTNum_Reset
                        TestID = TestID
                        CalDate = Now
                        Notes = ""
                        Workstation = GetComputerName()
                        TraceID1 = SQL.GetTraceID(Title, TestID)
                        TraceID = TraceID1
                    End If
                    ScanGPIB.GetTrace(Trace1Freq, IL1Data)
                    Trace1Freq = TrimX(Trace1Freq)
                    IL1Data_offs = TrimY(IL1Data, CDbl(frmMANUALTEST.txtOffset4.Text))
                    If TraceChecked And Not TweakMode Then
                        If UUTNum_Reset <= 5 Then
                            ReDim Preserve XArray(Trace1Freq.Count - 1)
                            ReDim Preserve YArray(Trace1Freq.Count - 1)
                            Array.Clear(YArray, 0, Trace1Freq.Count - 1)
                            XArray = Trace1Freq
                            YArray = IL1Data
                            SQL.SaveTrace(Title, TestID, TraceID)
                            YArray = IL1Data_offs
                            For y = 0 To YArray.Count - 1
                                AB_XArray(UUTNum_Reset - 1, y) = XArray(y)
                                AB1_YArray(UUTNum_Reset - 1, y) = YArray(y)
                            Next
                        End If
                    End If
                    MaxData(x) = MaxNoZero(IL1Data)
                End If
                If MutiCalChecked Then ScanGPIB.GetTraceMem()
                If VNAStr = "AG_E5071B" Then ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                If VNAStr = "N3383A" Then ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                If VNAStr <> "AG_E5071B" And VNAStr <> "N3383A" Then ScanGPIB.BusWrite("OPC?;DISPDDM;") 'Data/Memory
                Pts = Points
            Next x

            frmMANUALTEST.Refresh()
            For x = 2 To Ports
                If x = 2 Then
                    AB = MaxData(x)
                Else
                    If MaxData(x) > AB Then AB = MaxData(x)
                End If
            Next x

            If AB < Spec Then
                Return "Pass"
            Else
                Return "Fail"
            End If

            ABStr = Str(AB)
            ABStrArray = Split(AB, ".")
            If ABStrArray(0) = "0" Then GoTo Round
            AB = CDbl((ABStrArray(0) & "." & Left(ABStrArray(1), 3)))
Round:
            AB = AB + CDbl(frmMANUALTEST.txtOffset4.Text)
            AB = TruncateDecimal(AB, 2)
            If AB <= Spec Then
                AmplitudeBalanceCOMB = "Pass"
            Else
                AmplitudeBalanceCOMB = "Fail"
            End If
        End If
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function
    Public Function AmplitudeBalanceCOMB_Marker(TestRun As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 0) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceData(Points) As Object
        Dim TraceFreq(Points) As Object
        Dim MaxData(32) As Double
        Dim MinData(32) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace1Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim x As Long
        Dim t1 As Trace
        Dim ABMin As Double
        Dim ABMax As Double
        Dim ABArray(Points) As Double
        Dim i As Integer
        Dim Workstation As String
        Dim Title As String
        Dim NumPorts As Double
        Dim PortNum As Byte
        Dim Ports As Integer
        Dim ABStr As String = ""
        Dim ABStrArray(500) As String


        Title = ActiveTitle
        AmplitudeBalanceCOMB_Marker = ""
        ABSetDone = False
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset4.Text = "" Then frmMANUALTEST.txtOffset4.Text = 0
        t1 = New Trace
        Spec = GetSpecification("AmplitudeBalance")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset4.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Return "Pass"
            Else
                Return "Fail"
            End If
            frmMANUALTEST.Refresh()
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                ABTraceID1 = 4262
                ABTraceID2 = 4263
                GetTracePoints(ABTraceID1)

                ABArray = YArray

                ABMax = MaxNoZero(ABArray)
                i = MaxX(ABArray)

                GetTracePoints(ABTraceID2)
                ABArray = YArray
                ABMin = Min(YArray)

                AB = (ABMax - ABMin) / 2
                If AB < Spec Then
                    Return "Pass"
                Else
                    Return "Fail"
                End If

                ABStr = Str(AB)
                ABStrArray = Split(AB, ".")

                AB = CDbl((ABStrArray(0) & "." & Left(ABStrArray(1), 3)))
                AB = AB + CDbl(frmMANUALTEST.txtOffset4.Text)
                AB = TruncateDecimal(AB, 3)
            ElseIf PassChecked Then
                AB = Spec
                Return "Pass"
            ElseIf FailChecked Then
                AB = Spec + 10
                Return "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            NumPorts = GetSpecification("Ports")
            PortNum = CByte(NumPorts)
            Ports = Int(NumPorts)

            For x = 1 To Ports
                PortNum = CByte(x)
                ActiveTitle = "     TESTING AMPLITUDE BALANCE     SW POSITION " & x & "      "
                If SwitchedChecked Then  'Auto RF Switching
                    SetSwitchesPosition = PortNum
                    status = SwitchCom.SetSwitchPosition(PortNum) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                    frmMANUALTEST.Switch(x - 1)
                Else
                    MsgBox("Move Cables to RF Position " & x & " ")
                End If

                If MutiCalChecked Then
                    SetupVNA(True, 2)
                    If x <> 1 Then
                        If VNAStr = "AG_E5071B" Then
                            ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                            ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                            ScanGPIB.BusWrite(":INIT2:CONT ON") ' and start another sweep
                        ElseIf VNAStr = "N3383A" Then
                            ScanGPIB.BusWrite("CALC:PAR:SEL 'CH1_S21_1'")
                            ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                            ScanGPIB.BusWrite(":INIT1:CONT ON") ' and start another sweep
                        Else
                            ScanGPIB.BusWrite("OPC?;INPUDATA;") ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(gBuffer & ";")
                            ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
                            ScanGPIB.BusWrite("OPC?;CONT") ' and start another sweep
                        End If
                    End If
                End If
                If Not ABSetDone Or MutiCalChecked Then
                    If VNAStr = "AG_E5071B" Then
                        ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                        ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                    ElseIf VNAStr = "N3383A" Then
                        ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory Off"
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON")  ' Data On
                        ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                    Else
                        ScanGPIB.BusWrite("OPC?;DISPDATA;")
                        ScanGPIB.BusWrite("OPC?;CHAN2;")
                        ScanGPIB.BusWrite("OPC?;LOGM;")
                        ScanGPIB.BusWrite("OPC?;REFV 0")
                        ScanGPIB.BusWrite("OPC?;SCAL 10")
                    End If
                    ABSetDone = True
                End If
                If x = 1 Then If VNAStr = "AG_E5071B" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                If x = 1 Then If VNAStr = "N3383A" Then ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                If x = 1 Then If VNAStr <> "AG_E5071B" And VNAStr <> "N3383A" Then ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory

                frmMANUALTEST.Refresh()
                If MutiCalChecked Then ScanGPIB.GetTraceMem()
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                    IL1 = ScanGPIB.BusWrite(":CALC1:MARK1:Y?")  'Get Marker1 val
                    ScanGPIB.BusWrite(":CALC1:MARK2 ON")  'Marker2 on
                    ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MAX")  'Marker2 min
                    IL2 = ScanGPIB.BusWrite(":CALC1:MARK2:Y?")  'Get Marker2 val
                End If

                If VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                    IL1 = ScanGPIB.BusWrite(":CALC1:MARK1:Y?")  'Get Marker1 val
                    ScanGPIB.BusWrite(":CALC1:MARK2 ON")  'Marker2 on
                    ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MAX")  'Marker2 min
                    IL2 = ScanGPIB.DeviceQuery(":CALC1:MARK2:Y?")  'Get Marker2 val
                End If
                If VNAStr <> "AG_E5071B" And VNAStr <> "N3383A" Then
                    ScanGPIB.BusWrite("OPC?;DISPDDM;") 'Data/Memory
                    ScanGPIB.BusWrite("MARK1;")  'Marker 1 on
                    ScanGPIB.BusWrite("MARKMAXI;")  'Marker 1 max
                    IL1 = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker2 val
                    ScanGPIB.BusWrite("MARK2;")  'Marker2 on
                    ScanGPIB.BusWrite("MARKMINI;")  'Marker2 min
                    IL2 = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker2 val
                End If
            Next x

            frmMANUALTEST.Refresh()
            ABMax = Math.Abs(IL1)
            ABMin = Math.Abs(IL2)
            If ABMax > ABMin Then
                AB = ABMax
            Else
                AB = ABMin
            End If

            If AB < Spec Then
                Return "Pass"
            Else
                Return "Fail"
            End If
Round:
            AB = AB + CDbl(frmMANUALTEST.txtOffset4.Text)
            AB = TruncateDecimal(AB, 2)
            If AB <= Spec Then
                AmplitudeBalanceCOMB_Marker = "Pass"
            Else
                AmplitudeBalanceCOMB_Marker = "Fail"
            End If
        End If
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
        'Turn off all markers
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
            ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
            ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
        Else
            ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
        End If
    End Function

    Public Function Directivity(TestRun As String, Direction As Long, SpecType As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim x As Long
        Dim t1 As Trace
        Dim TraceID1 As Long
        Dim TraceID2 As Long
        Dim COUPArray(Points) As Double
        Dim i As Integer
        Dim Workstation As String
        Dim Title As String

        frmMANUALTEST.Refresh()
        Directivity = ""
        Title = ActiveTitle
        If Direction = 1 Then ActiveTitle = "     TESTING DIRECTIVITY   FORWARD DIRECTION    "
        If Direction = 2 Then ActiveTitle = "     TESTING DIRECTIVITY   REVERSE DIRECTION    "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset4.Text = "" Then frmMANUALTEST.txtOffset4.Text = 0
        t1 = New Trace
        Spec = GetSpecification("Directivity")
        SwitchPorts = SQL.GetSpecification("SwitchPorts")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset4.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Directivity = "Pass"
            Else
                Directivity = "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                COUPTraceID1 = 4262
                COUPTraceID2 = 4263
                GetTracePoints(COUPTraceID1)
                For i = 0 To Pts
                    COUPArray(i) = Math.Abs(YArray(i))
                Next

                COUPJ3 = MaxNoZero(COUPArray)
                i = MaxX(COUPArray)


                GetTracePoints(COUPTraceID2)
                COUPJ4 = Math.Abs(YArray(i))

                DIR = COUPJ4 - COUPJ3
                If DIR < Spec Then
                    Directivity = "Pass"
                Else
                    Directivity = "Fail"
                End If
            ElseIf PassChecked Then
                DIR = 17
                Directivity = "Pass"
                DIR = Spec + 10
                Directivity = "Fail"
            ElseIf FailChecked Then
                DIR = Spec + 10
                Directivity = "Fail"
            End If
        Else
            frmMANUALTEST.Refresh()
            If SwitchPorts = 1 Then ' ~~~~~~~ 6 port stuff ~~~~~~~~~
                If Direction = 1 Then
                    If SwitchedChecked Then   'Auto RF Switching
                        SetSwitchesPosition = 3
                        status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)

                        frmMANUALTEST.Switch(2)
                    Else
                        MsgBox("Move Cables to RF Position 3", vbOKOnly, "Manual Switch")
                    End If
                Else
                    If SwitchedChecked Then   'Auto RF Switching
                        SetSwitchesPosition = 5
                        status = SwitchCom.SetSwitchPosition(5) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)

                        frmMANUALTEST.Switch(4)
                    Else
                        MsgBox("Move Cables to RF Position 5", vbOKOnly, "Manual Switch")
                    End If
                End If
            Else ' ~~~~~~~ 4 port stuff ~~~~~~~~~
                If SwitchedChecked Then   'Auto RF Switching
                    SetSwitchesPosition = 3
                    status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)

                    frmMANUALTEST.Switch(2)
                Else
                    MsgBox("Move Cables to RF Position 3", vbOKOnly, "Manual Switch")
                End If
            End If

            If MutiCalChecked Then SetupVNA(True, 1)
            System.Threading.Thread.Sleep(500)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & 0 - GetSpecification("Coupling"))
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 5")
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & 0 - GetSpecification("Coupling"))
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 5")
            Else
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                ScanGPIB.BusWrite("OPC?;LOGM;")
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                ScanGPIB.BusWrite("OPC?;REFV " & 0 - GetSpecification("Coupling"))
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                ScanGPIB.BusWrite("OPC?;SCAL 5")
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
            End If
            ExtraAvg(2)
            frmMANUALTEST.Refresh()
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Coupling J3"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID1 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID1
            End If
            ScanGPIB.GetTrace(TraceFreq, COUP1DataDir)
            TraceFreq = TrimX(TraceFreq)
            COUP1Data_offs = TrimY(COUP1DataDir, CDbl(frmMANUALTEST.txtOffset4.Text))
            If TraceChecked And Not TweakMode Then
                If UUTNum_Reset <= 5 Then
                    ReDim Preserve XArray(TraceFreq.Count - 1)
                    ReDim Preserve YArray(TraceFreq.Count - 1)
                    Array.Clear(YArray, 0, TraceFreq.Count - 1)
                    XArray = TraceFreq
                    YArray = COUP1DataDir
                    SQL.SaveTrace(Title, TestID, TraceID)
                    YArray = COUP1Data_offs
                End If
            End If

            If SpecType = "SINGLE DIRECTIONAL COUPLER" Then
                If SwitchedChecked Then  'Auto RF Switching
                    SetSwitchesPosition = 3
                    status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                    frmMANUALTEST.Switch(2)
                Else
                    MsgBox("Move Cables to RF Position 3", vbOKOnly, "Manual Switch")
                End If

                If MutiCalChecked Then SetupVNA(True, 1)
            Else
                If SwitchPorts = 1 Then ' ~~~~~~~ 6 port stuff ~~~~~~~~~
                    If Direction = 1 Then
                        If SwitchedChecked Then  'Auto RF Switching
                            SetSwitchesPosition = 2
                            status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                            System.Threading.Thread.Sleep(500)
                            frmMANUALTEST.Switch(1)
                        Else
                            MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
                        End If
                    Else
                        If SwitchedChecked Then  'Auto RF Switching
                            SetSwitchesPosition = 4
                            status = SwitchCom.SetSwitchPosition(4) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                            System.Threading.Thread.Sleep(500)
                            frmMANUALTEST.Switch(3)
                        Else
                            MsgBox("Move Cables to RF Position 4", vbOKOnly, "Manual Switch")
                        End If
                    End If
                Else '~~~~~~~~~~~~~4 port stuff~~~~~~~~~~~~~~
                    If SwitchedChecked Then  'Auto RF Switching
                        SetSwitchesPosition = 2
                        status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)
                        frmMANUALTEST.Switch(1)
                    Else
                        MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
                    End If
                End If
                If MutiCalChecked Then SetupVNA(True, 2)
            End If
            frmMANUALTEST.Refresh()
            System.Threading.Thread.Sleep(500)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & 0 - (GetSpecification("Coupling") + GetSpecification("Directivity")))
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & 0 - (GetSpecification("Coupling") + GetSpecification("Directivity")))
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
            Else
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                ScanGPIB.BusWrite("OPC?;LOGM;")
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                ScanGPIB.BusWrite("OPC?;REFV " & 0 - (GetSpecification("Coupling") + GetSpecification("Directivity")))
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                ScanGPIB.BusWrite("OPC?;SCAL 10")
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
            End If
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Coupling J4"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID2 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID2
            End If
            ScanGPIB.GetTrace(TraceFreq, COUP2DataDir)
            TraceFreq = TrimX(TraceFreq)
            COUP2Data_offs = TrimY(COUP2DataDir, CDbl(frmMANUALTEST.txtOffset4.Text))
            If TraceChecked And Not TweakMode Then
                If UUTNum_Reset <= 5 Then
                    ReDim Preserve XArray(TraceFreq.Count - 1)
                    ReDim Preserve YArray(TraceFreq.Count - 1)
                    Array.Clear(YArray, 0, TraceFreq.Count - 1)
                    XArray = TraceFreq
                    YArray = COUP2DataDir
                    SQL.SaveTrace(Title, TestID, TraceID)
                    YArray = COUP2Data_offs
                    For x = 0 To YArray.Count - 1
                        COUP2_YArray(UUTNum_Reset - 1, x) = YArray(x)
                    Next
                End If
            End If

            'For x = 0 To 200
            '    If x = 0 Then COUPJ3 = Math.Abs(COUP1Data_offs(x))
            '    If x <> 0 Then COUPJ3 = COUPJ3 + Math.Abs(COUP1Data_offs(x))
            'Next x
            'COUPJ3 = COUPJ3 / 201

            COUPJ3 = Math.Abs(MaxNoZero(COUP1Data_offs))

            COUPJ4 = Math.Abs(MaxNoZero(COUP2Data_offs))

            If Direction = 1 Then
                ReturnVal1 = Math.Abs(COUPJ4 - COUPJ3)
            Else
                ReturnVal2 = Math.Abs(COUPJ4 - COUPJ3)
            End If

            If Direction = 1 Then
                DIR = ReturnVal1
            ElseIf SpecType = "SINGLE DIRECTIONAL COUPLER" Then
                DIR = Math.Abs(COuP - COUPJ3)
            Else
                If ReturnVal1 < ReturnVal2 Then
                    DIR = ReturnVal1
                Else
                    DIR = ReturnVal2
                End If
            End If

            DIR = Format(TruncateDecimal(DIR, 1) + frmMANUALTEST.txtOffset4.Text, "0.0")
            If DIR >= Spec Then
                Directivity = "Pass"
            Else
                Directivity = "Fail"
            End If
        End If

        ActiveTitle = Title
        frmMANUALTEST.Refresh()
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function
    Public Function Directivity_Marker(TestRun As String, Direction As Long, SpecType As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim t1 As Trace
        Dim COUPArray(Points) As Double
        Dim i As Integer
        Dim Workstation As String
        Dim Title As String
        frmMANUALTEST.Refresh()
        Directivity_Marker = ""
        Title = ActiveTitle
        If Direction = 1 Then ActiveTitle = "     TESTING DIRECTIVITY   FORWARD DIRECTION    "
        If Direction = 2 Then ActiveTitle = "     TESTING DIRECTIVITY   REVERSE DIRECTION    "
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset4.Text = "" Then frmMANUALTEST.txtOffset4.Text = 0
        t1 = New Trace
        Spec = GetSpecification("Directivity")
        SwitchPorts = SQL.GetSpecification("SwitchPorts")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset4.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Directivity_Marker = "Pass"
            Else
                Directivity_Marker = "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                COUPTraceID1 = 4262
                COUPTraceID2 = 4263
                GetTracePoints(COUPTraceID1)
                For i = 0 To Pts
                    COUPArray(i) = Math.Abs(YArray(i))
                Next

                COUPJ3 = MaxNoZero(COUPArray)
                i = MaxX(COUPArray)

                GetTracePoints(COUPTraceID2)
                COUPJ4 = Math.Abs(YArray(i))

                DIR = COUPJ4 - COUPJ3
                If DIR < Spec Then
                    Directivity_Marker = "Pass"
                Else
                    Directivity_Marker = "Fail"
                End If
            ElseIf PassChecked Then
                DIR = 17
                Directivity_Marker = "Pass"
                DIR = Spec + 10
                Directivity_Marker = "Fail"
            ElseIf FailChecked Then
                DIR = Spec + 10
                Directivity_Marker = "Fail"
            End If
        Else
            frmMANUALTEST.Refresh()
            If SwitchPorts = 1 Then ' ~~~~~~~ 6 port stuff ~~~~~~~~~
                If Direction = 1 Then
                    If SwitchedChecked Then   'Auto RF Switching
                        SetSwitchesPosition = 3
                        status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)

                        frmMANUALTEST.Switch(2)
                    Else
                        MsgBox("Move Cables to RF Position 3", vbOKOnly, "Manual Switch")
                    End If
                Else
                    If SwitchedChecked Then   'Auto RF Switching
                        SetSwitchesPosition = 5
                        status = SwitchCom.SetSwitchPosition(5) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)

                        frmMANUALTEST.Switch(4)
                    Else
                        MsgBox("Move Cables to RF Position 5", vbOKOnly, "Manual Switch")
                    End If
                End If
            Else ' ~~~~~~~ 4 port stuff ~~~~~~~~~
                If SwitchedChecked Then   'Auto RF Switching
                    SetSwitchesPosition = 3
                    status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)

                    frmMANUALTEST.Switch(2)
                Else
                    MsgBox("Move Cables to RF Position 3", vbOKOnly, "Manual Switch")
                End If
            End If

            If MutiCalChecked Then SetupVNA(True, 1)
            'System.Threading.Thread.Sleep(500)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & 0 - (GetSpecification("Coupling") + GetSpecification("Directivity")))
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                ExtraAvg(2)
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 min
                COUPJ1_Marker = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 min
                COUPJ1_Marker = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
                ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & 0 - (GetSpecification("Coupling") + GetSpecification("Directivity")))
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                ExtraAvg(2)
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 
                COUPJ1_Marker = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
            Else
                ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("OPC?;LOGM;")
                ScanGPIB.BusWrite("OPC?;REFV " & 0 - (GetSpecification("Coupling") + GetSpecification("Directivity")))
                ScanGPIB.BusWrite("OPC?;SCAL 10")
                ScanGPIB.BusWrite("MARK1;")  'Marker 1 on
                ExtraAvg(2)
                ScanGPIB.BusWrite("MARKMAXI;")  'Marker 1 max
                COUPJ1_Marker = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
            End If
            frmMANUALTEST.Refresh()


            If SpecType = "SINGLE DIRECTIONAL COUPLER" Then
                If SwitchedChecked Then   'Auto RF Switching
                    SetSwitchesPosition = 3
                    status = SwitchCom.SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                    frmMANUALTEST.Switch(2)
                Else
                    MsgBox("Move Cables to RF Position 3", vbOKOnly, "Manual Switch")
                End If

                If MutiCalChecked Then SetupVNA(True, 1)
            Else
                If SwitchPorts = 2 Then ' ~~~~~~~ 6 port stuff ~~~~~~~~~
                    If Direction = 1 Then
                        If SwitchedChecked Then  'Auto RF Switching
                            SetSwitchesPosition = 4
                            status = SwitchCom.SetSwitchPosition(4) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                            System.Threading.Thread.Sleep(500)
                            frmMANUALTEST.Switch(3)
                        Else
                            MsgBox("Move Cables to RF Position 4", vbOKOnly, "Manual Switch")
                        End If
                    Else
                        If SwitchedChecked Then  'Auto RF Switching
                            SetSwitchesPosition = 2
                            status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                            status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                            System.Threading.Thread.Sleep(500)
                            frmMANUALTEST.Switch(1)
                        Else
                            MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
                        End If
                    End If
                Else '~~~~~~~~~~~~~4 port stuff~~~~~~~~~~~~~~
                    If SwitchedChecked Then  'Auto RF Switching
                        SetSwitchesPosition = 2
                        status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)
                        frmMANUALTEST.Switch(1)
                    Else
                        MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
                    End If
                End If

                If MutiCalChecked Then SetupVNA(True, 2)
            End If
            frmMANUALTEST.Refresh()
            System.Threading.Thread.Sleep(500)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & 0 - GetSpecification("Coupling"))
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 5")
                COUPJ2_Marker = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & 0 - GetSpecification("Coupling"))
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 5")
                ExtraAvg(2)
                COUPJ2_Marker = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
            Else
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("OPC?;LOGM;")
                ScanGPIB.BusWrite("OPC?;REFV " & 0 - GetSpecification("Coupling"))
                ScanGPIB.BusWrite("OPC?;SCAL 5")
                ExtraAvg(2)
                Delay(500)
                COUPJ2_Marker = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
                Delay(500)
            End If


            COUPJ3 = Math.Abs(COUPJ1_Marker)
            COUPJ4 = Math.Abs(COUPJ2_Marker)
            If Direction = 1 Then
                ReturnVal1 = COUPJ3 - COUPJ4
            Else
                ReturnVal2 = COUPJ3 - COUPJ4
            End If

            If Direction = 1 Then
                DIR = ReturnVal1
            ElseIf SpecType = "SINGLE DIRECTIONAL COUPLER" Then
                DIR = Math.Abs(COuP - Math.Abs(COUPJ2_Marker))
            Else
                If ReturnVal1 < ReturnVal2 Then
                    DIR = ReturnVal1
                Else
                    DIR = ReturnVal2
                End If
            End If

            DIR = Math.Abs(DIR)
            DIR = Format(TruncateDecimal(DIR, 1) + frmMANUALTEST.txtOffset4.Text, "0.0")
            If DIR >= Spec Then
                Directivity_Marker = "Pass"
            Else
                Directivity_Marker = "Fail"
            End If
        End If

        ActiveTitle = Title
        frmMANUALTEST.Refresh()
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function
    Public Function CoupledFlatness(TestRun As String, Direction As Long, SpecType As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 0) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim t1 As Trace
        Dim TraceID1 As Long
        Dim SpecificationID As Long
        Dim COUPArray(Points) As Double
        Dim i As Integer
        Dim Workstation As String
        Dim Title As String
        Dim Direction1Value As Double
        Dim Direction2Value As Double

        Title = ActiveTitle
        If Direction = 1 Then ActiveTitle = "     TESTING COUPLING FLATNESS   FORWARD DIRECTION    "
        If Direction = 2 Then ActiveTitle = "     TESTING COUPLING FLATNESS   REVERSE DIRECTION    "
        Workstation = GetComputerName()
        SwitchPorts = SQL.GetSpecification("SwitchPorts")
        CoupledFlatness = ""
        Workstation = GetComputerName()
        If frmMANUALTEST.txtOffset4.Text = "" Then frmMANUALTEST.txtOffset4.Text = 0
        t1 = New Trace
        Spec = GetSpecification("CoupledFlatness")
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset4.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                CoupledFlatness = "Pass"
            Else
                CoupledFlatness = "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                Pts = Points
                COUPTraceID1 = 4262
                GetTracePoints(COUPTraceID1)
                For i = 0 To Pts - 1
                    COUP1Data(i) = YArray(i)
                Next

                CF = MaxNoZero(COUP1Data) - COUP1Data.Min
                If CF < Spec Then
                    CoupledFlatness = "Pass"
                Else
                    CoupledFlatness = "Fail"
                End If
            ElseIf PassChecked Then
                CF = Spec
                CoupledFlatness = "Pass"
            ElseIf FailChecked Then
                CF = Spec + 10
                CoupledFlatness = "Fail"
            End If
            frmMANUALTEST.Refresh()
        Else
            frmMANUALTEST.Refresh()
            If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then
                If SwitchedChecked Then  'Auto RF Switching
                    SetSwitchesPosition = 2
                    status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                    frmMANUALTEST.Switch(1)
                Else
                    MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
                End If
                If MutiCalChecked Then SetupVNA(True, 1)
            ElseIf SpecType = "DUAL DIRECTIONAL COUPLER" And Direction = 1 Then
                If SwitchPorts = 1 Then
                    If SwitchedChecked Then  'Auto RF Switching
                        SetSwitchesPosition = 2
                        status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)
                        frmMANUALTEST.Switch(1)
                    Else
                        MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
                    End If
                Else
                    If SwitchedChecked Then  'Auto RF Switching
                        SetSwitchesPosition = 2
                        status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)
                        frmMANUALTEST.Switch(1)
                    Else
                        MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
                    End If
                End If

                If MutiCalChecked Then SetupVNA(True, 1)
            ElseIf SpecType = "DUAL DIRECTIONAL COUPLER" And Direction = 2 Then
                If SwitchPorts = 1 Then
                    If SwitchedChecked Then  'Auto RF Switching
                        SetSwitchesPosition = 4
                        status = SwitchCom.SetSwitchPosition(4) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)
                        frmMANUALTEST.Switch(3)
                    Else
                        MsgBox("Move Cables to RF Position 4", vbOKOnly, "Manual Switch")
                    End If
                Else
                    If SwitchedChecked Then  'Auto RF Switching
                        SetSwitchesPosition = 2
                        status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                        status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                        System.Threading.Thread.Sleep(500)
                        frmMANUALTEST.Switch(1)
                    Else
                        MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
                    End If
                End If


                If MutiCalChecked Then SetupVNA(True, 3)
            End If
            frmMANUALTEST.Refresh()
            System.Threading.Thread.Sleep(100)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & 0 - GetSpecification("Coupling"))
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & 0 - GetSpecification("Coupling"))
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
            Else
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                ScanGPIB.BusWrite("OPC?;LOGM;")
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                ScanGPIB.BusWrite("OPC?;REFV " & 0 - GetSpecification("Coupling"))
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                ScanGPIB.BusWrite("OPC?;SCAL 10")
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
            End If
            If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or (SpecType = "DUAL DIRECTIONAL COUPLER" And Direction = 1) Then Title = "Manual " & TestRun & "Coupling Flatness J3"
            If SpecType = "DUAL DIRECTIONAL COUPLER" And Direction = 2 Then Title = "Manual " & TestRun & "Coupling Flatness J4"

            If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Then
                If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                    SerialNumber = "UUT1"
                    TestID = TestID
                    CalDate = Now
                    Notes = ""
                    Workstation = GetComputerName()
                    TraceID1 = SQL.GetTraceID(Title, TestID)
                    TraceID = TraceID1
                End If
                ScanGPIB.GetTrace(TraceFreq, COUP1FlatData)
                TraceFreq = TrimX(TraceFreq)
                COUP1FlatData_offs = TrimY(COUP1FlatData, CDbl(frmMANUALTEST.txtOffset4.Text))
                If TraceChecked And Not TweakMode Then
                    ReDim Preserve XArray(TraceFreq.Count - 1)
                    ReDim Preserve YArray(TraceFreq.Count - 1)
                    Array.Clear(YArray, 0, TraceFreq.Count - 1)
                    XArray = TraceFreq
                    YArray = COUP1FlatData
                    SQL.SaveTrace(Title, TestID, TraceID)
                End If
            Else
                If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                    SerialNumber = "UUT1"
                    TestID = TestID
                    CalDate = Now
                    Notes = ""
                    Workstation = GetComputerName()
                    TraceID1 = SQL.GetTraceID(Title, TestID)
                    TraceID = TraceID1
                End If
                ScanGPIB.GetTrace(TraceFreq, COUP2FlatData)
                TraceFreq = TrimX(TraceFreq)
                COUP2FlatData = TrimY(COUP2FlatData, CDbl(frmMANUALTEST.txtOffset4.Text))
                If TraceChecked And Not TweakMode Then
                    ReDim Preserve XArray(TraceFreq.Count - 1)
                    ReDim Preserve YArray(TraceFreq.Count - 1)
                    Array.Clear(YArray, 0, TraceFreq.Count - 1)
                    XArray = TraceFreq
                    YArray = COUP2FlatData
                    SQL.SaveTrace(Title, TestID, TraceID)
                End If
            End If
            frmMANUALTEST.Refresh()
            If TraceChecked And Not TweakMode Then
                t1 = New Trace
                If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or (SpecType = "DUAL DIRECTIONAL COUPLER" And Direction = 1) Then t1.Title(TraceID1, "Coupling Flatness J3")
                If SpecType = "DUAL DIRECTIONAL COUPLER" And Direction = 2 Then t1.Title(TraceID1, "Coupling Flatness J4")
                t1.SpecID(TraceID1, SpecificationID)
                t1.TestID(TraceID1, TestID)
                t1.CalDate(TraceID1, Now)
                t1.Workstation(TraceID1, Workstation)
                t1.Notes(TraceID1, "")
                t1.SaveTrace(TraceID1, True)
            End If


            If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Or Direction = 1 Then
                CF = MaxNoZero(COUP1FlatData) - COUP1FlatData.Min
                CF = TruncateDecimal(CF, 2) + frmMANUALTEST.txtOffset5.Text
                CF = TruncateDecimal(CF, 2)
                If CF < Spec * 2 Then
                    CoupledFlatness = "Pass"
                Else
                    CoupledFlatness = "Fail"
                End If
                CF = (CF / 2)
                CF = TruncateDecimal(CF, 2)
            Else
                Direction1Value = MaxNoZero(COUP1FlatData) - COUP1FlatData.Min
                Direction2Value = MaxNoZero(COUP2FlatData) - COUP2FlatData.Min
                If Direction1Value > Direction2Value Then
                    CF = Direction1Value
                Else
                    CF = Direction2Value
                End If
                CF = TruncateDecimal(CF, 2) + frmMANUALTEST.txtOffset5.Text
                If CF < Spec * 2 Then
                    CoupledFlatness = "Pass"
                Else
                    CoupledFlatness = "Fail"
                End If
                CF = (CF / 2)
                CF = TruncateDecimal(CF, 2)
            End If
        End If
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
            ScanGPIB.BusWrite(":CALC1:FORM MLOG")
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
            ScanGPIB.BusWrite(":CALC1:FORM MLOG")
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
        Else
            ScanGPIB.BusWrite("OPC?;CHAN2;")
            ScanGPIB.BusWrite("OPC?;LOGM;")
            ScanGPIB.BusWrite("OPC?;REFV 0")
            ScanGPIB.BusWrite("OPC?;SCAL 10")
        End If
        frmMANUALTEST.Refresh()
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function

    Public Function PhaseBalance(TestRun As String, SpecType As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim TraceID2 As Long
        Dim TraceID3 As Long
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace1Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim i As Integer
        Dim data(201) As Double
        Dim Workstation As String
        Dim ABArray(Points) As Double
        Dim Title As String

        Title = ActiveTitle
        ActiveTitle = "     TESTING PHASE BALANCE   SW POSITION 1     "
        PhaseBalance = ""
        Workstation = GetComputerName()
        Spec = GetSpecification("PhaseBalance")
        If frmMANUALTEST.txtOffset5.Text = "" Then frmMANUALTEST.txtOffset5.Text = 0
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset5.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                PhaseBalance = "Pass"
            Else
                PhaseBalance = "Fail"
            End If
            frmMANUALTEST.Refresh()
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4266
                TraceID2 = 4267
                GetTracePoints(TraceID1)
                Trace1Data = YArray
                Trace1Freq = XArray
                GetTracePoints(TraceID2)
                Trace2Data = YArray
                Trace2Freq = XArray
                For i = 0 To Pts - 1
                    ABArray(i) = Math.Abs(Math.Abs(Trace1Data(i) - Trace2Data(i)) - 90)
                Next

                PB = MaxNoZero(ABArray)
                PB = PB + CDbl(frmMANUALTEST.txtOffset5.Text)


                If PB < Spec Then
                    PhaseBalance = "Pass"
                Else
                    PhaseBalance = "Fail"
                End If
            ElseIf PassChecked Then
                PB = Spec
                PhaseBalance = "Pass"
            ElseIf frmMANUALTEST.Fail.Checked Then
                PB = Spec + 10
                PhaseBalance = "Fail"
            End If
        Else
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 1 Then
                    SetSwitchesPosition = 1
                    status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(0)
            Else
                MsgBox("Move Cables to RF Position 1", vbOKOnly, "Manual Switch")
                If MutiCalChecked Then RecallCal(1)
            End If

            If MutiCalChecked Then SetupVNA(True, 1)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
                ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
            Else
                ScanGPIB.BusWrite("OPC?;DISPDATA;")
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("OPC?;PHAS;")
                ScanGPIB.BusWrite("OPC?;REFV 0")
                ScanGPIB.BusWrite("OPC?;SCAL 10")
            End If
            ExtraAvg()
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Phase Balance J3"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID1 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID1
            End If
            ScanGPIB.GetTrace(Trace1Freq, Trace1Data)
            Trace1Freq = TrimX(Trace1Freq)
            Trace1Data_offs = TrimY(Trace1Data, CDbl(frmMANUALTEST.txtOffset5.Text))
            If TraceChecked And Not TweakMode Then
                If UUTNum_Reset <= 5 Then
                    ReDim Preserve XArray(TraceFreq.Count - 1)
                    ReDim Preserve YArray(TraceFreq.Count - 1)
                    Array.Clear(YArray, 0, TraceFreq.Count - 1)
                    XArray = Trace1Freq
                    YArray = Trace1Data
                    SQL.SaveTrace(Title, TestID, TraceID)
                    YArray = Trace1Data_offs
                End If
            End If

            frmMANUALTEST.Refresh()
            System.Threading.Thread.Sleep(500)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            Else
                ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
            End If
            If MutiCalChecked Then ScanGPIB.GetTraceMem()
            ActiveTitle = "     TESTING PHASE BALANCE   SW POSITION 2       "

            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 2 Then
                    SetSwitchesPosition = 2
                    status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(1)
            Else
                MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
            End If
            frmMANUALTEST.Refresh()

            If MutiCalChecked Then
                SetupVNA(True, 2)
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                    ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(":INIT1:CONT ON") ' and start another sweep
                    ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                    If SpecType = "90 DEGREE COUPLER" Then
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -90")
                    ElseIf SpecType.Contains("BALUN") Then
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -180")
                    End If
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF") 'Memory On"
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                    ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(":INIT1:CONT ON") ' and start another sweep
                    ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                    If SpecType = "90 DEGREE COUPLER" Then
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -90")
                    ElseIf SpecType.Contains("BALUN") Then
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -180")
                    End If
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory On"
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON")  ' Data On
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;PHAS;")
                    ScanGPIB.BusWrite("OPC?;REFV 0")
                    ScanGPIB.BusWrite("OPC?;SCAL 10")
                    ScanGPIB.BusWrite("OPC?;INPUDATA," & gBuffer & ";") ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(gBuffer & ";")
                    If SpecType = "90 DEGREE COUPLER" Then
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -90")
                    ElseIf SpecType.Contains("BALUN") Then
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -180")
                    End If
                    ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
                    ScanGPIB.BusWrite("OPC?;CONT") ' and start another sweep
                End If
            End If
            ExtraAvg()
            If (VNAStr = "AG_E5071B") And Not MutiCalChecked Then
                If MutiCalChecked = 0 Then
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM ON")  'Data On
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                End If
                If SpecType = "90 DEGREE COUPLER" Then
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -90")
                ElseIf SpecType.Contains("BALUN") Then
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -180")
                End If

            ElseIf (VNAStr = "N3383A") And Not MutiCalChecked Then
                If MutiCalChecked = 0 Then
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM ON") 'Data On
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
                End If
                If SpecType = "90 DEGREE COUPLER" Then
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -90")
                ElseIf SpecType.Contains("BALUN") Then
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -180")
                End If
            Else
                If MutiCalChecked = 0 Then ScanGPIB.BusWrite("OPC?;DISPDATM;") 'Data and Memory
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                If SpecType = "90 DEGREE COUPLER" Then ScanGPIB.BusWrite("OPC?;REFV -90")
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                If SpecType = "90 DEGREE COUPLER" Then
                    ScanGPIB.BusWrite("OPC?;REFV 90")
                ElseIf SpecType.Contains("BALUN") Then
                    ScanGPIB.BusWrite("OPC?;REFV -180")
                End If
            End If

            If (VNAStr = "AG_E5071B") And Not MutiCalChecked Then
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM ON")  'Data On
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
            ElseIf (VNAStr = "N3383A") And Not MutiCalChecked Then
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM ON")  'Data On
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
            Else
                ScanGPIB.BusWrite("OPC?;DISPDATM;") 'Data and Memory
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
            End If
            ExtraAvg()
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Phase Balance J4"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID2 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID2
            End If
            ScanGPIB.GetTrace(Trace2Freq, Trace2Data)
            Trace2Freq = TrimX(Trace2Freq)
            Trace2Data_offs = TrimY(Trace2Data, CDbl(frmMANUALTEST.txtOffset5.Text))
            If TraceChecked And Not TweakMode Then
                ReDim Preserve XArray(TraceFreq.Count - 1)
                ReDim Preserve YArray(TraceFreq.Count - 1)
                Array.Clear(YArray, 0, TraceFreq.Count - 1)
                XArray = Trace2Freq
                YArray = Trace2Data
                SQL.SaveTrace(Title, TestID, TraceID)
                YArray = Trace2Data_offs
            End If
            System.Threading.Thread.Sleep(500)
            If VNAStr = "AG_E5071B" Then
                If Not MutiCalChecked Then ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
            ElseIf VNAStr = "N3383A" Then
                If Not MutiCalChecked Then ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
            Else
                If Not MutiCalChecked Then ScanGPIB.BusWrite("OPC?;DISPDDM;") 'Data/Memory
                If Not MutiCalChecked And VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
            End If

            frmMANUALTEST.Refresh()
            If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                Title = "Manual " & TestRun & "Phase Balance D/M"
                SerialNumber = "UUT" & UUTNum_Reset
                TestID = TestID
                CalDate = Now
                Notes = ""
                Workstation = GetComputerName()
                TraceID3 = SQL.GetTraceID(Title, TestID)
                TraceID = TraceID3
            End If
            ScanGPIB.GetTrace(TraceFreq, TraceData)
            TraceFreq = TrimX(TraceFreq)
            TraceData_offs = TrimY(TraceData, CDbl(frmMANUALTEST.txtOffset5.Text))
            If TraceChecked And Not TweakMode Then
                ReDim Preserve XArray(TraceFreq.Count - 1)
                ReDim Preserve YArray(TraceFreq.Count - 1)
                Array.Clear(YArray, 0, TraceFreq.Count - 1)
                XArray = TraceFreq
                YArray = TraceData
                SQL.SaveTrace(Title, TestID, TraceID)
                YArray = TraceData_offs
                If UUTNum_Reset <= 5 Then
                    For x = 0 To YArray.Count - 1
                        PB1_YArray(UUTNum_Reset - 1, x) = YArray(x)
                    Next
                End If
            End If

            If SpecType = "90 DEGREE COUPLER" Then
                For i = 0 To Pts - 1
                    ABArray(i) = Math.Abs(90 - Math.Abs(TraceData(i)))
                Next
            Else
                For i = 0 To Pts - 1
                    ABArray(i) = Math.Abs(180 - Math.Abs(TraceData(i)))
                Next
            End If

            PB = MaxNoZero(ABArray)
            PB = TruncateDecimal(PB, 1)
            PB = Format(PB + CDbl(frmMANUALTEST.txtOffset5.Text), "0.0")
            System.Threading.Thread.Sleep(1000)


            ' Put Back to IL so user can have a reference
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                Delay(100)
                ScanGPIB.BusWrite(":CALC1:MATH:FUNC NORM")
                ScanGPIB.BusWrite(":CALC1:MATH:FUNC NORM")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON")  ' Data On
                ScanGPIB.BusWrite(":CALC1:MATH:FUNC NORM")
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
            Else
                ScanGPIB.BusWrite("OPC?;DISPDATA;")
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("OPC?;LOGM;")
                ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
                ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("AmplitudeBalance"))
                ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
            End If
            PB = TruncateDecimal(PB, 1)
            ILSetDone = True
            If PB <= Spec Then
                PhaseBalance = "Pass"
            Else
                PhaseBalance = "Fail"
            End If
        End If
        ActiveTitle = Title
        frmMANUALTEST.Refresh()
        '************DEBUG CODE FOR JEN'S WORKSTATION*******************
        If VNAStr = "AG_E5071B" And User = "JEN" Then
            'SetSwitchesPosition = 1
            'status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            'frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
        Else
            SetSwitchesPosition = 1
            status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
            frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
        End If
    End Function

    Public Function PhaseBalance_Marker(TestRun As String, SpecType As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim TraceID2 As Long
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace1Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim i As Integer
        Dim data(201) As Double
        Dim Workstation As String
        Dim ABArray(Points) As Double
        Dim Title As String

        Title = ActiveTitle
        ActiveTitle = "     TESTING PHASE BALANCE   SW POSITION 1     "
        PhaseBalance_Marker = ""
        Workstation = GetComputerName()
        Spec = GetSpecification("PhaseBalance")
        If frmMANUALTEST.txtOffset5.Text = "" Then frmMANUALTEST.txtOffset5.Text = 0
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset5.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                PhaseBalance_Marker = "Pass"
            Else
                PhaseBalance_Marker = "Fail"
            End If
            frmMANUALTEST.Refresh()
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4266
                TraceID2 = 4267
                GetTracePoints(TraceID1)
                Trace1Data = YArray
                Trace1Freq = XArray
                GetTracePoints(TraceID2)
                Trace2Data = YArray
                Trace2Freq = XArray
                For i = 0 To Pts - 1
                    ABArray(i) = Math.Abs(Math.Abs(Trace1Data(i) - Trace2Data(i)) - 90)
                Next

                PB = MaxNoZero(ABArray)
                PB = PB + CDbl(frmMANUALTEST.txtOffset5.Text)


                If PB < Spec Then
                    PhaseBalance_Marker = "Pass"
                Else
                    PhaseBalance_Marker = "Fail"
                End If
            ElseIf PassChecked Then
                PB = Spec
                PhaseBalance_Marker = "Pass"
            ElseIf frmMANUALTEST.Fail.Checked Then
                PB = Spec + 10
                PhaseBalance_Marker = "Fail"
            End If
        Else
            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 1 Then
                    SetSwitchesPosition = 1
                    status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                End If
                frmMANUALTEST.Switch(0)
            Else
                MsgBox("Move Cables to RF Position 1", vbOKOnly, "Manual Switch")
                If MutiCalChecked Then RecallCal(1)
            End If

            If MutiCalChecked Then SetupVNA(True, 1)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
                ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
            Else
                ScanGPIB.BusWrite("OPC?;DISPDATA;")
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("OPC?;PHAS;")
                ScanGPIB.BusWrite("OPC?;REFV 0")
                ScanGPIB.BusWrite("OPC?;SCAL 10")
            End If
            frmMANUALTEST.Refresh()
            'System.Threading.Thread.Sleep(500)
            ExtraAvg()
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
            Else
                ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
            End If
            If MutiCalChecked Then ScanGPIB.GetTraceMem()
            ActiveTitle = "     TESTING PHASE BALANCE   SW POSITION 2       "

            If SwitchedChecked Then  'Auto RF Switching
                If SetSwitchesPosition <> 2 Then
                    SetSwitchesPosition = 2
                    status = SwitchCom.SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(1000)
                End If
                frmMANUALTEST.Switch(1)
            Else
                MsgBox("Move Cables to RF Position 2", vbOKOnly, "Manual Switch")
            End If
            frmMANUALTEST.Refresh()

            If MutiCalChecked Then
                SetupVNA(True, 2)
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                    ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                    ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(":INIT1:CONT ON") ' and start another sweep
                    ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                    If SpecType = "90 DEGREE COUPLER" Then ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -90")
                    If SpecType.Contains("BALUN") Then ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -180")
                    Delay(100)
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF") 'Memory On"
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                    ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                    ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(":INIT1:CONT ON") ' and start another sweep
                    ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                    If SpecType = "90 DEGREE COUPLER" Then ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV -90")
                    If SpecType.Contains("BALUN") Then ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV -180")
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory On"
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON")  ' Data On
                Else
                    ScanGPIB.BusWrite("OPC?;CHAN2;")
                    ScanGPIB.BusWrite("OPC?;PHAS;")
                    ScanGPIB.BusWrite("OPC?;REFV 0")
                    ScanGPIB.BusWrite("OPC?;SCAL 10")
                    ScanGPIB.BusWrite("OPC?;INPUDATA," & gBuffer & ";") ' Input Trace1 data to VNA Memory
                    ScanGPIB.BusWrite(gBuffer & ";")
                    ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
                    ScanGPIB.BusWrite("OPC?;CONT") ' and start another sweep
                End If
            End If
            ExtraAvg()
            If (VNAStr = "AG_E5071B") And Not MutiCalChecked Then
                If MutiCalChecked = 0 Then
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM ON")  'Data On
                    ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                End If
                If SpecType = "90 DEGREE COUPLER" Then ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -90")
                If SpecType.Contains("BALUN") Then ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV -180")
            ElseIf (VNAStr = "N3383A") And Not MutiCalChecked Then
                If MutiCalChecked = 0 Then
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM ON") 'Data On
                    ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
                End If
                If SpecType = "90 DEGREE COUPLER" Then ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV -90")
                If SpecType.Contains("BALUN") Then ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV -180")
            Else
                If MutiCalChecked = 0 Then ScanGPIB.BusWrite("OPC?;DISPDATM;") 'Data and Memory
                If SpecType = "90 DEGREE COUPLER" Then ScanGPIB.BusWrite("OPC?;REFV -90")
                If SpecType.Contains("BALUN") Then ScanGPIB.BusWrite("OPC?;REFV -180")
            End If

            If (VNAStr = "AG_E5071B") And Not MutiCalChecked Then
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM ON")  'Data On
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
            ElseIf (VNAStr = "N3383A") And Not MutiCalChecked Then
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM ON")  'Data On
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
            Else
                ScanGPIB.BusWrite("OPC?;DISPDATM;") 'Data and Memory
                'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
            End If

            If VNAStr = "AG_E5071B" Then
                If Not MutiCalChecked Then ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                ExtraAvg()
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                Delay(100)
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                PB1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                ScanGPIB.BusWrite(":CALC1:MARK2 ON")  'Marker2 on
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MIN")  'Marker2 min
                Delay(100)
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MIN")  'Marker2 min
                PB2 = ScanGPIB.MarkerQuery(":CALC1:MARK2:Y?")  'Get Marker2 val
            ElseIf VNAStr = "N3383A" Then
                If Not MutiCalChecked Then ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                ExtraAvg()
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                Delay(50)
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                PB1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                ScanGPIB.BusWrite(":CALC1:MARK2 ON")  'Marker2 on
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MIN")  'Marker2 min
                Delay(50)
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MIN")  'Marker2 min
                PB2 = ScanGPIB.MarkerQuery(":CALC1:MARK2:Y?")  'Get Marker2 val
            Else
                If Not MutiCalChecked Then ScanGPIB.BusWrite("OPC?;DISPDDM;") 'Data/Memory
                If Not MutiCalChecked And VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
                ScanGPIB.BusWrite("MARK1;")  'Marker1 on
                ExtraAvg()
                Delay(50)
                ScanGPIB.BusWrite("MARKMAXI;")  'Marker1 max
                PB1 = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
                ScanGPIB.BusWrite("MARK2;")  'Marker2 on
                ScanGPIB.BusWrite("MARKMINI;")  'Marker2 max
                PB2 = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker2 val
            End If

            frmMANUALTEST.Refresh()

            If SpecType = "90 DEGREE COUPLER" Then
                PB1 = Math.Abs(90 - Math.Abs(PB1))
                PB2 = Math.Abs(90 - Math.Abs(PB2))
            Else
                PB1 = Math.Abs(180 - Math.Abs(PB1))
                PB2 = Math.Abs(180 - Math.Abs(PB2))
            End If

            If Math.Abs(PB1) > Math.Abs(PB2) Then
                PB = PB1
            Else
                PB = PB2
            End If
            PB = Math.Abs(TruncateDecimal(PB, 1))
            PB = Format(PB + CDbl(frmMANUALTEST.txtOffset5.Text), "0.0")

            ' Put Back to IL so user can have a reference
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                ScanGPIB.BusWrite(":CALC1:MATH:FUNC NORM")
                ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON")  ' Data On
                ScanGPIB.BusWrite(":CALC1:MATH:FUNC NORM")
                ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
            Else
                ScanGPIB.BusWrite("OPC?;DISPDATA;")
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("OPC?;LOGM;")
                ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
                ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("AmplitudeBalance"))
                ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
            End If
            PB = TruncateDecimal(PB, 1)
            ILSetDone = True
            If PB <= Spec Then
                PhaseBalance_Marker = "Pass"
            Else
                PhaseBalance_Marker = "Fail"
            End If
        End If

        ActiveTitle = Title
        frmMANUALTEST.Refresh()
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"

        'Turn off all markers
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
            ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
            ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
        Else
            ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
        End If
    End Function
    Public Function PhaseBalanceCOMB(TestRun As String, SpecType As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim TraceID2 As Long
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace1Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim MaxData(32) As Double
        Dim MinData(32) As Double
        Dim x As Long
        Dim i As Integer
        Dim data(201) As Double
        Dim Workstation As String
        Dim ABArray(Points) As Double
        Dim Title As String
        Dim NumPorts As Double
        Dim PortNum As Byte
        Dim Ports As Integer

        Title = ActiveTitle
        PhaseBalanceCOMB = ""
        PBSetDone = False
        Workstation = GetComputerName()
        Spec = GetSpecification("PhaseBalance")
        If frmMANUALTEST.txtOffset5.Text = "" Then frmMANUALTEST.txtOffset5.Text = 0
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset5.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Return "Pass"
            Else
                Return "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4266
                TraceID2 = 4267
                GetTracePoints(TraceID1)
                Trace1Data = YArray
                Trace1Freq = XArray
                GetTracePoints(TraceID2)
                Trace2Data = YArray
                Trace2Freq = XArray
                For i = 0 To Pts - 1
                    ABArray(i) = Math.Abs(Math.Abs(Trace1Data(i) - Trace2Data(i)) - 90)
                Next

                PB = MaxNoZero(ABArray)
                PB = PB + CDbl(frmMANUALTEST.txtOffset5.Text)

                If PB < Spec Then
                    Return "Pass"
                Else
                    Return "Fail"
                End If
            ElseIf PassChecked Then
                PB = Spec
                Return "Pass"
            ElseIf FailChecked Then
                PB = Spec + 10
                Return "Fail"
            End If
        Else
            frmMANUALTEST.Refresh()
            NumPorts = GetSpecification("Ports")
            PortNum = CByte(NumPorts)
            Ports = Int(NumPorts)
            PBSetDone = False
            For x = 1 To Ports
                PortNum = CByte(x)
                ActiveTitle = "     TESTING PHASE BALANCE     SW POSITION " & x & "      "
                If SwitchedChecked Then  'Auto RF Switching
                    SetSwitchesPosition = PortNum
                    status = SwitchCom.SetSwitchPosition(PortNum) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                    frmMANUALTEST.Switch(x - 1)
                Else
                    MsgBox("Move Cables to RF Position " & x & " ")
                End If
                If x = 1 Then
                    If VNAStr = "AG_E5071B" Then
                        ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                        ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                        ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                    ElseIf VNAStr = "N3383A" Then
                        ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory Off"
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
                        ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                        ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                    Else
                        ScanGPIB.BusWrite("OPC?;DISPDATA;")
                        ScanGPIB.BusWrite("OPC?;CHAN2;")
                        ScanGPIB.BusWrite("OPC?;S21;")
                        ScanGPIB.BusWrite("OPC?;PHAS;")
                        ScanGPIB.BusWrite("OPC?;REFV 0")
                        ScanGPIB.BusWrite("OPC?;SCAL 10")
                        ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
                    End If
                End If
                If MutiCalChecked And x = 1 Then ScanGPIB.GetTraceMem()
                If MutiCalChecked Then
                    SetupVNA(True, x)
                    If x <> 1 Then
                        If VNAStr = "AG_E5071B" Then
                            ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                            ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                            ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                            ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                        ElseIf VNAStr = "N3383A" Then
                            ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory Off"
                            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
                            ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                            ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                            ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                        Else
                            ScanGPIB.BusWrite("OPC?;DISPDATA;")
                            ScanGPIB.BusWrite("OPC?;CHAN2;")
                            ScanGPIB.BusWrite("OPC?;S21;")
                            ScanGPIB.BusWrite("OPC?;PHAS;")
                            ScanGPIB.BusWrite("OPC?;REFV 0")
                            ScanGPIB.BusWrite("OPC?;SCAL 10")
                            ScanGPIB.BusWrite("OPC?;INPUDATA;") ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(gBuffer & ";")
                            ScanGPIB.BusWrite("OPC?;CONT") ' and start another sweep
                        End If
                    End If
                End If
                ExtraAvg(2)
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                Else
                    ScanGPIB.BusWrite("OPC?;DISPDDM;") 'Data/Memory
                End If

                If TraceChecked And Not TweakMode And frmMANUALTEST.SaveData Then ' Database Trace Data
                    Title = "Manual " & TestRun & "PhaseBalance Port " & PortNum
                    SerialNumber = "UUT" & UUTNum_Reset
                    TestID = TestID
                    CalDate = Now
                    Notes = ""
                    Workstation = GetComputerName()
                    TraceID1 = SQL.GetTraceID(Title, TestID)
                    TraceID = TraceID1
                End If
                ScanGPIB.GetTrace(Trace1Freq, IL1Data)
                Trace1Freq = TrimX(Trace1Freq)
                IL1Data_offs = TrimY(IL1Data, CDbl(frmMANUALTEST.txtOffset5.Text))
                If TraceChecked And Not TweakMode Then
                    If UUTNum_Reset <= 5 Then
                        ReDim Preserve XArray(TraceFreq.Count - 1)
                        ReDim Preserve YArray(TraceFreq.Count - 1)
                        Array.Clear(YArray, 0, TraceFreq.Count - 1)
                        XArray = Trace1Freq
                        YArray = IL1Data
                        SQL.SaveTrace(Title, TestID, TraceID)
                        YArray = IL1Data_offs
                        For y = 0 To YArray.Count - 1
                            PB_XArray(UUTNum_Reset - 1, y) = XArray(y)
                            PB1_YArray(UUTNum_Reset - 1, y) = YArray(y)
                        Next
                    End If
                End If
                ReDim Preserve MaxData(x - 1)
                MaxData(x - 1) = MaxNoZero(IL1Data)

                If MutiCalChecked Then ScanGPIB.GetTraceMem()

                Pts = Points
            Next x
            ExtraAvg(2)
            frmMANUALTEST.Refresh()
            For x = 0 To Ports - 1
                If x = 2 Then
                    PB = MaxData(x)
                Else
                    If MaxData(x) > PB Then PB = MaxData(x)
                End If
            Next x

            PB = PB + CDbl(frmMANUALTEST.txtOffset5.Text)
            System.Threading.Thread.Sleep(500)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF") 'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                ScanGPIB.BusWrite(":CALC1:MATH:FUNC NORM")
                ScanGPIB.BusWrite(":CALC1:MATH:FUNC NORM")
                ' Put Back to IL so user can have a reference
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory Off"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
                ScanGPIB.BusWrite(":CALC1:MATH:FUNC NORM")
                ' Put Back to IL so user can have a reference
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                frmMANUALTEST.Refresh()
            Else
                ScanGPIB.BusWrite("OPC?;DISPDATA;")
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("OPC?;LOGM;")
                ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
                ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("AmplitudeBalance"))
            End If

            frmMANUALTEST.Refresh()
            ILSetDone = True

            PB = TruncateDecimal(PB, 1)
            If PB <= Spec Then
                PhaseBalanceCOMB = "Pass"
            Else
                PhaseBalanceCOMB = "Fail"
            End If
        End If
        frmMANUALTEST.Refresh()
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
    End Function

    Public Function PhaseBalanceCOMB_Marker(TestRun As String, SpecType As String, Optional ResumeTesting As Boolean = False, Optional TestID As Long = 1) As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Spec As Double
        Dim TraceID1 As Long
        Dim TraceID2 As Long
        Dim TraceData(Points) As Double
        Dim TraceFreq(Points) As Double
        Dim Trace1Data(Points) As Double
        Dim Trace1Freq(Points) As Double
        Dim Trace2Data(Points) As Double
        Dim Trace2Freq(Points) As Double
        Dim MaxData(32) As Double
        Dim MinData(32) As Double
        Dim x As Long
        Dim i As Integer
        Dim data(201) As Double
        Dim Workstation As String
        Dim ABArray(Points) As Double
        Dim Title As String
        Dim NumPorts As Double
        Dim PortNum As Byte
        Dim Ports As Integer

        Title = ActiveTitle
        PhaseBalanceCOMB_Marker = ""
        PBSetDone = False
        Workstation = GetComputerName()
        Spec = GetSpecification("PhaseBalance")
        If frmMANUALTEST.txtOffset5.Text = "" Then frmMANUALTEST.txtOffset5.Text = 0
        If ResumeTesting Then
            RetrnVal = RetrnVal + CDbl(frmMANUALTEST.txtOffset5.Text)
            If TruncateDecimal(RetrnVal, 2) <= Spec Then
                Return "Pass"
            Else
                Return "Fail"
            End If
        ElseIf Debug Then  ' Simulated Data
            If DBDataChecked Then
                TraceID1 = 4266
                TraceID2 = 4267
                GetTracePoints(TraceID1)
                Trace1Data = YArray
                Trace1Freq = XArray
                GetTracePoints(TraceID2)
                Trace2Data = YArray
                Trace2Freq = XArray
                For i = 0 To Pts - 1
                    ABArray(i) = Math.Abs(Math.Abs(Trace1Data(i) - Trace2Data(i)) - 90)
                Next

                PB = MaxNoZero(ABArray)
                PB = PB + CDbl(frmMANUALTEST.txtOffset5.Text)


                If PB < Spec Then
                    Return "Pass"
                Else
                    Return "Fail"
                End If
            ElseIf PassChecked Then
                PB = Spec
                Return "Pass"
            ElseIf FailChecked Then
                PB = Spec + 10
                Return "Fail"
            End If
        Else
            frmMANUALTEST.Refresh()
            NumPorts = GetSpecification("Ports")
            PortNum = CByte(NumPorts)
            Ports = Int(NumPorts)
            PBSetDone = False
            For x = 1 To Ports
                PortNum = CByte(x)
                ActiveTitle = "     TESTING PHASE BALANCE     SW POSITION " & x & "      "
                If SwitchedChecked Then  'Auto RF Switching
                    SetSwitchesPosition = PortNum
                    status = SwitchCom.SetSwitchPosition(PortNum) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
                    status = SwitchCom.GetSwitchPosition(StatusRet) ' note Status Return in Binary
                    System.Threading.Thread.Sleep(500)
                    frmMANUALTEST.Switch(x - 1)
                Else
                    MsgBox("Move Cables to RF Position " & x & " ")
                End If
                If x = 1 Then
                    If VNAStr = "AG_E5071B" Then
                        ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF")  'Memory Off"
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                        ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV 10")
                        ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                    ElseIf VNAStr = "N3383A" Then
                        ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory Off"
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
                        ScanGPIB.BusWrite(":CALC1:FORM PHAS")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV 0")
                        ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV 10")
                        ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                    Else
                        ScanGPIB.BusWrite("OPC?;DISPDATA;")
                        ScanGPIB.BusWrite("OPC?;CHAN2;")
                        ScanGPIB.BusWrite("OPC?;S21;")
                        ScanGPIB.BusWrite("OPC?;PHAS;")
                        ScanGPIB.BusWrite("OPC?;REFV 0")
                        ScanGPIB.BusWrite("OPC?;SCAL 10")
                        ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
                    End If
                End If

                If MutiCalChecked Then
                    SetupVNA(True, 2)
                    If x <> 1 Then
                        If VNAStr = "AG_E5071B" Then
                            ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                            ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                            ScanGPIB.BusWrite(":INIT2:CONT ON") ' and start another sweep
                        ElseIf VNAStr = "N3383A" Then
                            ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                            ScanGPIB.BusWrite(":CALC1:DATA:SMEM " & gBuffer) ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(":CALC1:MATH:MEM") 'Data into Memory
                            ScanGPIB.BusWrite(":INIT1:CONT ON") ' and start another sweep
                        Else
                            ScanGPIB.BusWrite("OPC?;INPUDATA;") ' Input Trace1 data to VNA Memory
                            ScanGPIB.BusWrite(gBuffer & ";")
                            ScanGPIB.BusWrite("OPC?;DATI;") 'Data into Memory
                            ScanGPIB.BusWrite("OPC?;CONT") ' and start another sweep
                        End If
                    End If
                End If
                If MutiCalChecked And x = 1 Then ScanGPIB.GetTraceMem()
                If VNAStr = "AG_E5071B" Then
                    ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ExtraAvg()
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                    PB1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                    ScanGPIB.BusWrite(":CALC1:MARK2 ON")  'Marker2 on
                    ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                    ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MIN")  'Marker2 min
                    PB2 = ScanGPIB.MarkerQuery(":CALC1:MARK2:Y?")  'Get Marker2 val
                ElseIf VNAStr = "N3383A" Then
                    ScanGPIB.BusWrite(":CALC1:MATH:FUNC DIV") 'Data/Memory
                    ScanGPIB.BusWrite(":CALC1:MARK1 ON")  'Marker 1 on
                    ExtraAvg()
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:EXEC")
                    ScanGPIB.BusWrite(":CALC1:MARK1:FUNC:TYPE MAX")  'Marker 1 max
                    PB1 = ScanGPIB.MarkerQuery(":CALC1:MARK1:Y?")  'Get Marker1 val
                    ScanGPIB.BusWrite(":CALC1:MARK2 ON")  'Marker2 on
                    ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:EXEC")
                    ScanGPIB.BusWrite(":CALC1:MARK2:FUNC:TYPE MIN")  'Marker2 min
                    PB2 = ScanGPIB.MarkerQuery(":CALC1:MARK2:Y?")  'Get Marker2 val
                Else
                    ScanGPIB.BusWrite("OPC?;DISPDDM;") 'Data/Memory
                    ScanGPIB.BusWrite("MARK1;")  'Marker 1 on
                    ExtraAvg(2)
                    Delay(500)
                    ScanGPIB.BusWrite("MARKMAXI;")  'Marker 1 max
                    PB1 = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker1 val
                    ScanGPIB.BusWrite("MARK2;")  'Marker2 on
                    ScanGPIB.BusWrite("MARKMINI;")  'Marker2 min
                    PB2 = ScanGPIB.MarkerQuery("OUTPMARK;")  'Get Marker2 val
                    If PB1 > PB2 Then
                        PB = PB1
                    Else
                        PB = PB2
                    End If
                    ReDim Preserve MaxData(x - 1)
                    MaxData(x - 1) = PB
                End If
                Pts = Points
            Next x

            frmMANUALTEST.Refresh()
            PB = MaxNoZero(MaxData)
            PB = TruncateDecimal(PB, 1)
            PB = Format(PB + CDbl(frmMANUALTEST.txtOffset5.Text), "0.0")
            'System.Threading.Thread.Sleep(500)
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:MEM OFF") 'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:STAT ON") ' Data On
                ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                ' Put Back to IL so user can have a reference
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:RLEV " & GetLoss())
                ScanGPIB.BusWrite(":DISP:WIND1:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:MEM OFF") 'Memory On"
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:STAT ON") ' Data On
                ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
                ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
                ' Put Back to IL so user can have a reference
                ScanGPIB.BusWrite(":CALC1:FORM MLOG")
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:RLEV " & GetLoss())
                ScanGPIB.BusWrite(":DISP:WIND2:TRAC2:Y:PDIV " & GetSpecification("AmplitudeBalance"))
                frmMANUALTEST.Refresh()
            Else
                ScanGPIB.BusWrite("OPC?;DISPDATA;")
                ' Put Back to IL so user can have a reference
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("OPC?;LOGM;")
                ScanGPIB.BusWrite("OPC?;REFV " & GetLoss())
                ScanGPIB.BusWrite("OPC?;SCAL " & GetSpecification("AmplitudeBalance"))
                ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
            End If

            frmMANUALTEST.Refresh()
            ILSetDone = True

            PB = TruncateDecimal(PB, 1)
            If PB <= Spec Then
                PhaseBalanceCOMB_Marker = "Pass"
            Else
                PhaseBalanceCOMB_Marker = "Fail"
            End If
        End If

        frmMANUALTEST.Refresh()
        ActiveTitle = Title
        SetSwitchesPosition = 1
        status = SwitchCom.SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
        'Turn off all markers
        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
            ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite(":CALC1:MARK1 OFF")  'Marker1 off
            ScanGPIB.BusWrite(":CALC1:MARK2 OFF")  'Marker2 off
            ScanGPIB.BusWrite(":CALC1:MARK3 OFF")  'Marker3 off
        Else
            ScanGPIB.BusWrite("MARKOFF;")  'All Markers Off
        End If
    End Function
    Public Function InitializeSwitch(retSwitchNumber As Integer, retSerialList As String, retFirmware As Integer) As String
        Dim ConnectionGood As Integer
        Try
            retSwitchNumber = GetNumberOfSwitches()
            retSerialList = GetSNlist()
            retFirmware = GetFirmware()
            Dim model = GetSwitchModel()


            ConnectionGood = SwitchCom.Connect ' Note requires a few milliseconds to connect
            System.Threading.Thread.Sleep(1000)

            InitializeSwitch = SwitchCom.Get24VConnection
            SwitchModel = GetSwitchModel()
            If SwitchModel = "RC-1SP6T-A12" Then
                frmMANUALTEST.btSwitch1.Visible = True
                frmMANUALTEST.btSwitch2.Visible = True
                frmMANUALTEST.btSwitch3.Visible = True
                frmMANUALTEST.btSwitch4.Visible = True
                frmMANUALTEST.btSwitch5.Visible = True
                frmMANUALTEST.btSwitch6.Visible = True
                frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
            Else
                frmMANUALTEST.btSwitch1.Visible = True
                frmMANUALTEST.btSwitch2.Visible = True
                frmMANUALTEST.btSwitch3.Visible = True
                frmMANUALTEST.btSwitch4.Visible = True
                frmMANUALTEST.cmbSwitch.Text = "Switch POS 1"
            End If
        Catch ex As Exception

        End Try

    End Function

    Public Sub ExtraAvg(Optional CH As Integer = 2)
        If frmMANUALTEST.ExtraAverage.Checked Then
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite("SENS:AVER:COUNT " & CStr(frmMANUALTEST.AvgS.Text))
                ScanGPIB.BusWrite("SENS:AVER ON")
                ScanGPIB.BusWrite("SENS2:AVER:COUNT " & CStr(frmMANUALTEST.AvgS.Text))
                ScanGPIB.BusWrite("SENS2:AVER ON")
                Delay(2000)
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("SENS:AVER:COUNT " & CStr(frmMANUALTEST.AvgS.Text))
                ScanGPIB.BusWrite("SENS:AVER ON")
                ScanGPIB.BusWrite("SENS2:AVER:COUNT " & CStr(frmMANUALTEST.AvgS.Text))
                ScanGPIB.BusWrite("SENS2:AVER ON")
                Delay(2000)
            Else
                ScanGPIB.BusWrite("OPC?;CHAN1;")
                ScanGPIB.BusWrite("AVERFACT " & CStr(frmMANUALTEST.AvgS.Text))
                ScanGPIB.BusWrite("AVERO ON")
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("AVERFACT " & CStr(frmMANUALTEST.AvgS.Text))
                ScanGPIB.BusWrite("OPC?;AVERO ON")
                If CH = 1 Then ScanGPIB.BusWrite("OPC?;CHAN1;")
                Delay(2000)

            End If
        End If
    End Sub


    'Public Function PrintReport()
    'On Error GoTo ErrHandler
    'Dim SQLstr As String
    'Dim Test As String
    'Dim EXCEL As XLReport
    'Dim ATS As ADODB.Recordset
    'Dim Spec As ADODB.Recordset
    'Dim Directional As Boolean
    'Dim Row As Long
    'Dim TopFolder As String
    'Dim SubFolder As String
    'Dim t1 As Trace
    'Dim t2 As Trace
    'Dim t3 As Trace
    'Dim TraceID1 As Long
    'Dim TraceID2 As Long
    'Dim TraceID3 As Long
    '
    'SQLstr = "SELECT * from Specifications where JobNumber = '" & frmMANUALTEST.cmbJob.Text & "'"
    'Set Spec = New ADODB.Recordset
    'Spec.Open SQLstr, NetConn
    'If Spec.EOF Then Exit Function
    '
    'Set EXCEL = New XLReport
    'EXCEL.OpenTemplate ExcelTemplatePath & "TraceData.xls"
    'EXCEL.WriteToCell "C5", frmMANUALTEST.SerialNum.Text
    'EXCEL.WriteToCell "F2", Spec!JobNumber
    'EXCEL.WriteToCell "F3", Spec!PartNumber
    'EXCEL.WriteToCell "F4", Spec!Title
    '
    '    Row = 10
    '   If chkSignal(0) Then
    '        EXCEL.WriteToCell "A8", txtTitle1.Text
    '        Set t1 = New Trace
    '        If txtNet.Caption = "Network" Then TraceID1 = t1.GetTraceIDByTitleNet(txtTitle1, frmMANUALTEST.SerialNum, Me.cmbJob)
    '        If txtNet.Caption = "Local" Then TraceID1 = t1.GetTraceIDByTitleLocal(txtTitle1, frmMANUALTEST.SerialNum, Me.cmbJob)
    '        LoadTraceInfo TraceID1, 0
    '    End If
    '    If chkSignal(1) Then
    '        EXCEL.WriteToCell "C8", txtTitle2.Text
    '        Set t2 = New Trace
    '        If txtNet.Caption = "Network" Then TraceID2 = t2.GetTraceIDByTitleNet(txtTitle2, frmMANUALTEST.SerialNum, Me.cmbJob)
    '        If txtNet.Caption = "Local" Then TraceID2 = t2.GetTraceIDByTitleLocal(txtTitle2, frmMANUALTEST.SerialNum, Me.cmbJob)
    '        LoadTraceInfo TraceID2, 1
    '    End If
    '    If chkSignal(2) Then
    '        EXCEL.WriteToCell "E8", txtTitle3.Text
    '        Set t3 = New Trace
    '        If txtNet.Caption = "Network" Then TraceID3 = t3.GetTraceIDByTitleNet(txtTitle3, frmMANUALTEST.SerialNum, Me.cmbJob)
    '        If txtNet.Caption = "Local" Then TraceID3 = t3.GetTraceIDByTitleLocal(txtTitle3, frmMANUALTEST.SerialNum, Me.cmbJob)
    '        LoadTraceInfo TraceID3, 2
    '    End If
    '
    '    If TraceID1 = -1 Then
    '        MsgBox "No trace available"
    '       Exit Function
    '    End If
    '
    '    If txtNet.Caption = "Network" Then
    '        For n = 0 To (t1.GetPointsNet(TraceID1) - 1)
    '            If TraceID1 > 0 Then
    '                EXCEL.WriteToCell "A" & Row, t1.GetXDataNet(TraceID1, n)
    '                EXCEL.WriteToCell "B" & Row, t1.GetYDataNet(TraceID1, n)
    '            End If
    '            If TraceID2 > 0 Then
    '                EXCEL.WriteToCell "C" & Row, t2.GetXDataNet(TraceID2, n)
    '                EXCEL.WriteToCell "D" & Row, t2.GetYDataNet(TraceID2, n)
    '            End If
    '            If TraceID3 > 0 Then
    '                EXCEL.WriteToCell "E" & Row, t3.GetXDataNet(TraceID3, n)
    '                EXCEL.WriteToCell "F" & Row, t3.GetYDataNet(TraceID3, n)
    '            End If
    '            Row = Row + 1
    '        Next n
    '    Else
    '        For n = 0 To (t1.GetPointsLocal(TraceID1) - 1)
    '            If TraceID1 > 0 Then
    '                EXCEL.WriteToCell "A" & Row, t1.GetXDataLocal(TraceID1, n)
    '                EXCEL.WriteToCell "B" & Row, t1.GetYDataLocal(TraceID1, n)
    '            End If
    '            If TraceID2 > 0 Then
    '                EXCEL.WriteToCell "C" & Row, t2.GetXDataLocal(TraceID2, n)
    '                EXCEL.WriteToCell "D" & Row, t2.GetYDataLocal(TraceID2, n)
    '            End If
    '            If TraceID3 > 0 Then
    '                EXCEL.WriteToCell "E" & Row, t3.GetXDataLocal(TraceID3, n)
    '                EXCEL.WriteToCell "F" & Row, t3.GetYDataLocal(TraceID3, n)
    '            End If
    '            Row = Row + 1
    '        Next n
    '    End If
    '
    '    If Spec!SpecType = "90 DEGREE COUPLER" Then TopFolder = "90_Degree\"
    '    If Spec!SpecType = "90 DEGREE COUPLER SMD" Then TopFolder = "90_Degree_SMD\"
    '    If InStr(Spec!SpecType, "DIRECTIONAL COUPLER") Then TopFolder = "Directional_Couplers\"
    '    If InStr(Spec!SpecType, "DIRECTIONAL COUPLER SMD") Then TopFolder = "Directional_Couplers_SMD\"
    '    If Spec!SpecType.Contains("COMBINER/DIVIDER") Then TopFolder = "Combiner-Divider\"
    '    If Spec!SpecType = "COMBINER/DIVIDER SMD" Then TopFolder = "Combiner-Divider_SMD\"
    '
    '     'Network Save
    '
    '    If Dir$(TestDataPath, vbDirectory) = "" Then MkDir (TestDataPath)
    '    If Dir$(TestDataPath & TopFolder, vbDirectory) = "" Then MkDir (TestDataPath & TopFolder)
    '    SubFolder = TestDataPath & TopFolder & frmMANUALTEST.cmbPart.Text & "-" & frmMANUALTEST.cmbJob.Text & "\"
    '    If Dir$(SubFolder, vbDirectory) = "" Then MkDir (SubFolder)
    '    EXCEL.SaveAs SubFolder, "TraceData.xls"
    '
    '
    'Exit Function
    '
    'ErrHandler:
    ' MsgBox "An Error has Occured In The Form_Load() Procedure" & vbCr & "Report This Error To R_K_T_ASHOKA@RediffMail.com" & vbCr & "Error Details :-" & vbCr & "Error Number : " & Err.Number & vbCr & "Error Description : " & Err.Description, vbCritical, "FlexGrid Example"
    '
    '
    'End Function



End Module
