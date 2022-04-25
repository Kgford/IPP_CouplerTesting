Module MCC
    Public iniPathName As String = "C:/GIZMO/config_files/myconfig.ini"

    Public Sub LoadminiLAB()
        Dim BoardNum As Integer
        Dim typeVal As Integer
        Dim MyMessage As String
        Dim r As Microsoft.VisualBasic.MsgBoxResult

        BoardNum = 0
        Daqboard = New MccDaq.MccBoard(BoardNum)   '<======this is the default board number
        'change it to what InstaCal has assigned for your miniLAB-1008

        ulstat = Daqboard.BoardConfig.GetBoardType(typeVal)  ' Get the typeVal property from the MccBoard object
        If typeVal <> 0 Then
            If typeVal <> 117 Then  'Code for miniLAB-1008 is 117
                MyMessage = "A miniLAB-1008 was not assigned to Board " & BoardNum & " in InstaCal."
                r = MYMsgBox(MyMessage, vbExclamation, "miniLAB-1008 not detected.")
                End
            End If
        End If
    End Sub
    Public Function GetROBOTMovingSignal() As Boolean
        Try
            GetROBOTMovingSignal = 0
            If RobotStatus Then
                Dim BitValue As MccDaq.DigitalLogicState
                Dim PortType As MccDaq.DigitalPortType
                Dim BitNum As Integer
                PortType = MccDaq.DigitalPortType.AuxPort

                BitNum = 2
                ulstat = Daqboard.DBitIn(PortType, BitNum, BitValue)
                If ulstat.Value <> MccDaq.ErrorInfo.ErrorCode.NoErrors Then Stop

                If BitValue = MccDaq.DigitalLogicState.High Then
                    GetROBOTMovingSignal = True
                Else
                    GetROBOTMovingSignal = False
                End If
            End If
        Catch ex As Exception
            GetROBOTMovingSignal = 0
        End Try
    End Function
    Public Function GetROBOTErrorSignal() As Boolean
        Try
            GetROBOTErrorSignal = False
            If RobotStatus Then
                Dim BitValue As MccDaq.DigitalLogicState
                Dim PortType As MccDaq.DigitalPortType
                Dim BitNum As Integer
                PortType = MccDaq.DigitalPortType.AuxPort

                BitNum = 1
                ulstat = Daqboard.DBitIn(PortType, BitNum, BitValue)
                If ulstat.Value <> MccDaq.ErrorInfo.ErrorCode.NoErrors Then Stop

                If BitValue = MccDaq.DigitalLogicState.High Then
                    GetROBOTErrorSignal = True
                Else
                    GetROBOTErrorSignal = False
                End If
            End If
        Catch ex As Exception
            GetROBOTErrorSignal = True
        End Try

    End Function
    Public Function GetReadySignal() As Boolean
        Try
            GetReadySignal = False
            If RobotStatus Then
                Dim BitValue As MccDaq.DigitalLogicState
                Dim PortType As MccDaq.DigitalPortType
                Dim BitNum As Integer
                PortType = MccDaq.DigitalPortType.AuxPort

                BitNum = 0
                ulstat = Daqboard.DBitIn(PortType, BitNum, BitValue)
                If ulstat.Value <> MccDaq.ErrorInfo.ErrorCode.NoErrors Then Stop

                If BitValue = MccDaq.DigitalLogicState.High Then
                    GetReadySignal = True
                Else
                    GetReadySignal = False
                End If
            End If
        Catch ex As Exception
            GetReadySignal = 0
        End Try
    End Function
    Public Function GetInBits() As Integer()
        Try
            GetInBits = New Integer() {0, 1}
            If RobotStatus Then
                Dim BitValue As MccDaq.DigitalLogicState
                Dim PortType As MccDaq.DigitalPortType
                Dim BitNum As Integer
                Dim InBits(5) As Integer
                PortType = MccDaq.DigitalPortType.AuxPort
                For BitNum = 0 To 4
                    ulstat = Daqboard.DBitIn(PortType, BitNum, BitValue)
                    'If ulstat.Value <> MccDaq.ErrorInfo.ErrorCode.NoErrors Then Stop
                    If BitValue = MccDaq.DigitalLogicState.High Then
                        InBits(BitNum) = 1
                    Else
                        InBits(BitNum) = 0
                    End If
                Next
                GetInBits = InBits
            End If
        Catch ex As Exception
            ReadySignal = 0
            GetInBits = New Integer() {0, 1}
        End Try
    End Function

    Public Sub TestCompleteSignal(TrueFalse As Boolean)
        If RobotStatus Then
            Dim Range As MccDaq.Range
            Dim DataValue As Short
            Dim EngUnits As Single
            Dim IsValidNumber As Boolean
            Dim chan As Integer

            IsValidNumber = True
            ' send the digital output value to D/A 0 with MccDaq.MccBoard.AOut()

            'Try-Catch is a method of testing a value that is SUPPOSED TO BE GOOD.
            'If it is not, then the error is handled in the Catch portion of the code.
            Try
                If TrueFalse Then
                    EngUnits = Single.Parse(3.5)
                    TestComplete = True
                Else
                    EngUnits = Single.Parse(0)
                    TestComplete = False
                End If

            Catch ex As Exception
                MessageBox.Show(EngUnits + " is not a valid voltage value", "Invalid Voltage ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                IsValidNumber = False
            End Try

            If (IsValidNumber) Then 'if this is true then OK to convert the value and output a value

                Range = MccDaq.Range.Uni5Volts
                ulstat = Daqboard.FromEngUnits(Range, EngUnits, DataValue)
                If ulstat.Value <> MccDaq.ErrorInfo.ErrorCode.NoErrors Then Stop

                ' Parameters:
                '   Chan       :the D/A output channel
                '   Range      :ignored if board does not have programmable rage
                '   DataValue  :the value to send to Chan (comes from the FromEngUnits method)
                chan = 1  ' output channel
                Range = MccDaq.Range.Uni5Volts
                ulstat = Daqboard.AOut(chan, Range, DataValue)
                If ulstat.Value <> MccDaq.ErrorInfo.ErrorCode.NoErrors Then Stop

                'this corrects the number entered in the text box if it is beyond the scale of the
                'miniLAB-1008' DAC range of 0 to 5. If someone types in 6 (for example) it 
                'calculates the maximum value available based upon the converted value from the
                'FromEngUnits method.  Normally you would not need this.
                EngUnits = DataValue * 0.00488
                If ulstat.Value <> MccDaq.ErrorInfo.ErrorCode.NoErrors Then Stop
            End If
        End If

    End Sub

    Public Sub TestRunningSignal(TrueFalse As Boolean)
        If RobotStatus Then
            Dim Range As MccDaq.Range
            Dim DataValue As Short
            Dim EngUnits As Single
            Dim IsValidNumber As Boolean
            Dim chan As Integer

            IsValidNumber = True
            ' send the digital output value to D/A 0 with MccDaq.MccBoard.AOut()

            'Try-Catch is a method of testing a value that is SUPPOSED TO BE GOOD.
            'If it is not, then the error is handled in the Catch portion of the code.
            Try
                If TrueFalse Then
                    EngUnits = Single.Parse(3.5)
                Else
                    EngUnits = Single.Parse(0)
                End If

            Catch ex As Exception
                MessageBox.Show(EngUnits + " is not a valid voltage value", "Invalid Voltage ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                IsValidNumber = False
            End Try

            If (IsValidNumber) Then 'if this is true then OK to convert the value and output a value

                Range = MccDaq.Range.Uni5Volts
                ulstat = Daqboard.FromEngUnits(Range, EngUnits, DataValue)
                If ulstat.Value <> MccDaq.ErrorInfo.ErrorCode.NoErrors Then Stop

                ' Parameters:
                '   Chan       :the D/A output channel
                '   Range      :ignored if board does not have programmable rage
                '   DataValue  :the value to send to Chan (comes from the FromEngUnits method)
                chan = 0  ' output channel
                Range = MccDaq.Range.Uni5Volts
                ulstat = Daqboard.AOut(chan, Range, DataValue)
                If ulstat.Value <> MccDaq.ErrorInfo.ErrorCode.NoErrors Then Stop

                'this corrects the number entered in the text box if it is beyond the scale of the
                'miniLAB-1008' DAC range of 0 to 5. If someone types in 6 (for example) it 
                'calculates the maximum value available based upon the converted value from the
                'FromEngUnits method.  Normally you would not need this.
                EngUnits = DataValue * 0.00488
                If ulstat.Value <> MccDaq.ErrorInfo.ErrorCode.NoErrors Then Stop
            End If
        End If
    End Sub


    Public Function GetConfigurationVal(ByVal FileName As String, ByVal SectionName As String, ByVal KeyName As String) As String

        Dim sectionfound As Boolean = False
        Dim pn As String = ""
        If RobotStatus Then
            If IO.File.Exists(FileName) Then
                For Each s As String In IO.File.ReadAllLines(FileName)
                    If sectionfound And s.StartsWith("[") Then Exit For
                    If s.ToLower = "[" & SectionName.ToLower & "]" Then sectionfound = True
                    If sectionfound And s.ToLower.StartsWith(KeyName.ToLower) Then
                        Dim str() As String = s.Split(":"c)
                        If str.Length > 1 Then
                            pn = str(1)
                            Exit For
                        End If
                    End If
                Next
            End If
        End If
        Return pn
    End Function

    Public Sub saveConfigurationVal(ByVal FileName As String, ByVal KeyName As String, ByVal value As String)
        Dim x As Integer
        Dim str(10) As String
        If RobotStatus Then
            If IO.File.Exists(FileName) Then
                For Each s As String In IO.File.ReadAllLines(FileName)
                    ReDim Preserve str(x)
                    str(x) = s
                    If str(x).Contains(KeyName) Then
                        str(x) = KeyName & ": " & value
                    End If
                    x += 1
                Next
            End If
            IO.File.WriteAllLines(FileName, str)
        End If
    End Sub
    Public Sub PassFailSignal(PassFail As String)
        If RobotStatus Then
            Try
                saveConfigurationVal(iniPathName, "test_results", PassFail)
            Catch ex As Exception
                MessageBox.Show(" PassFail parameter not saved", "Can't save file ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Public Function GetRetest() As Boolean
        GetRetest = False
        If RobotStatus Then
            Try
                GetRetest = CType(GetConfigurationVal(iniPathName, "retestTF", "retest"), Boolean)
            Catch ex As Exception
                GetRetest = False
                ' MessageBox.Show(" Retest parameter not sretrieved", "Can't retrieve data ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Function

    Public Sub GetTrayStatus()
        If RobotStatus Then
            Try
                NewTray = GetConfigurationVal(iniPathName, "traystatus", "new_tray")
                PassTray = GetConfigurationVal(iniPathName, "traystatus", "pass_tray")
                FailTray = GetConfigurationVal(iniPathName, "traystatus", "fail_tray")
            Catch ex As Exception

                MessageBox.Show(" Retest parameter not sretrieved", "Can't retrieve data ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

End Module
