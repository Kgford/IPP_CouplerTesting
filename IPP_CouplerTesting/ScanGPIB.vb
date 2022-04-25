
Imports System.IO
Module ScanGPIB
    Public myVNA As DirectIO.DirectIO
    Public BussAddress As Integer = 7
    Public Function connect(ByVal address As String, ByVal timeout As Integer) As Boolean
        Try
            myVNA = New DirectIO.DirectIO(address, False, 5000)
            myVNA.Timeout(timeout)
            'myVNA.WriteLine("FORM REAL,32")
            'myVNA.WriteLine("FORM ASC")
            'myVNA.WriteLine("SWE:POIN 601")

        Catch ex As Exception
            MYMsgBox("VNA Error: " + ex.Message)
            Return False
        End Try
        Return True
    End Function

    Public Function BusWrite(cmdStr As String) As Boolean
        Try
            myVNA.WriteLine(cmdStr)
        Catch ex As Exception
            'MYMsgBox("VNA Error: " + ex.Message)
            Return False
        End Try
        Return True
    End Function

    Public Function BusRead() As String
        Try
            BusRead = myVNA.Read()
        Catch ex As Exception
            'MYMsgBox("VNA Error: " + ex.Message)
            BusRead = "Error"
        End Try
    End Function

    Public Function DeviceQuery(cmdStr As String, Optional Trace As Boolean = False) As String

        Try
            'send command string
            BusWrite(cmdStr)
            If Trace Then System.Threading.Thread.Sleep(300)
            System.Threading.Thread.Sleep(50)
            DeviceQuery = BusRead()
        Catch ex As Exception
            MYMsgBox("VNA Error: " + ex.Message)
            DeviceQuery = "Error"
        End Try

    End Function

    Public Function MarkerQuery(cmdStr As String) As String
        Dim tempval As String
        Dim valArray(3) As String
        Try
            'send command string
            BusWrite(cmdStr)
            System.Threading.Thread.Sleep(50)
            tempval = BusRead()
            valArray = Split(tempval, ",")
            MarkerQuery = valArray(0)

        Catch ex As Exception
            MYMsgBox("VNA Error: " + ex.Message)
            MarkerQuery = "Error"
        End Try

    End Function

    Public Function GetStartFreq() As Double
        Dim cmdStr As String
        On Error GoTo Trap


        If VNAStr = "AG_E5071B" Or VNAStr = "N3383A" Then
            cmdStr = DeviceQuery("SENS:FREQ:STAR?")
        Else
            cmdStr = DeviceQuery("STAR?")
            System.Threading.Thread.Sleep(2000)
            cmdStr = DeviceQuery("STAR?")
        End If
        GetStartFreq = CDbl(cmdStr)
        Exit Function
Trap:
        'MYMsgBox "GPIB Communication lost. Please restart and resume", , "GPIB Failure Get Start Freq"
        '
    End Function


    Public Function GetStopFreq() As Double
        Dim cmdStr As String

        On Error GoTo Trap

        If VNAStr = "AG_E5071B" Or VNAStr = "N3383A" Then
            cmdStr = DeviceQuery("SENS:FREQ:STOP?")
        Else
            cmdStr = DeviceQuery("STOP?")
            System.Threading.Thread.Sleep(2000)
            cmdStr = DeviceQuery("STOP?")
        End If
        GetStopFreq = CDbl(cmdStr)
        Exit Function
Trap:
        'MYMsgBox "GPIB Communication lost. Please restart and resume", , " Get Stop Freq"
        'End
    End Function

    Public Function GetNumPoints() As Double
        Dim cmdStr As String
        On Error GoTo Trap
        If VNAStr = "AG_E5071B" Or VNAStr = "N3383A" Then
            cmdStr = DeviceQuery("SENS:SWE:POIN?")
        Else
            cmdStr = DeviceQuery("POIN?")
        End If
        GetNumPoints = CDbl(cmdStr) * 100

        Exit Function
Trap:
        GetNumPoints = 201
        'MYMsgBox "GPIB Communication lost. Please restart and resume", , "GPIB Failure  Get Points"
        'End
    End Function

    Public Function GetModel() As String
        Dim cmdStr As String
        On Error GoTo Trap

        If Debug Then
            GetModel = "HP_8753E"
            Exit Function
        End If
        System.Threading.Thread.Sleep(1000)
        cmdStr = DeviceQuery("*IDN?")

        GetModel = ""
        If InStr(cmdStr, "E5071B") Then GetModel = "AG_E5071B"
        If InStr(cmdStr, "8720ES") Then GetModel = "HP_8720ES"
        If InStr(cmdStr, "8753E") Then GetModel = "HP_8753E"
        If InStr(cmdStr, "8753C") Then GetModel = "HP_8753C"

        If GetModel = "" Then GetModel = VNAStr ' Fail safe
        Exit Function
Trap:
        'MYMsgBox "GPIB Communication lost. Please restart and resume", , "GPIB Failure  Get Points"
        'End
    End Function

    Public Function GetTrace(rtnX() As Double, rtnY() As Double, Optional startf As Double = 0, Optional stopf As Double = 0) As Long
        Dim Buffer As String = ""
        Dim Buffer2 As String = ""
        Dim i As Long
        Dim freq As Double
        Dim tmp As String
        Dim ypt As Double
        Dim t1 As New Trace
        Dim TraceID As Long
        Dim x As Integer
        Array.Clear(rtnY, 0, Pts)
        Array.Clear(rtnX, 0, Pts)

        i = 0
        Pts = 201
        If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(1000)

        If EAveraging Then
            Averages = frmAUTOTEST.AvgS.Text
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite("SENS:AVER OFF")
                ScanGPIB.BusWrite("SENS:AVER ON")
                ScanGPIB.BusWrite("SENS2:AVER OFF")
                ScanGPIB.BusWrite("SENS2:AVER ON")
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("SENS:AVER OFF")
                ScanGPIB.BusWrite("SENS:AVER ON")
                ScanGPIB.BusWrite("SENS2:AVER OFF")
                ScanGPIB.BusWrite("SENS2:AVER ON")
            Else
                ScanGPIB.BusWrite("OPC?;CHAN1")
                ScanGPIB.BusWrite("AVERO OFF")
                ScanGPIB.BusWrite("AVERO ON")
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("AVERO OFF")
                ScanGPIB.BusWrite("AVERO ON")
            End If
            System.Threading.Thread.Sleep(Averages * 500)
        End If
        GetTrace = 0
        'If TraceChecked And Not TweakMode Then
        '    TraceID = t1.OpenTrace
        '    GetTrace = TraceID
        'End If
        If VNAStr = "HP_8720ES" Then Pts = Pts - 1
        ' tell the vna to send the ascii data
        If VNAStr = "AG_E5071B" Then
            Buffer = DeviceQuery(":CALC1:DATA:FDAT?", True)
            Do While i < Pts
                If startf = 0 Then
                    If freq = 0 Then
                        freq = Star
                    Else
                        freq = Star + i * ((Sto - Star) / (Pts - 1))
                    End If
                Else
                    If freq = 0 Then
                        freq = startf
                    Else
                        freq = startf + i * ((stopf - startf) / (Pts - 1))
                    End If

                End If
                ' get the corresponding y points
                tmp = Mid(Buffer, i * 40 + 1, 40)
                ypt = CDbl(Val(tmp))
                ' add the point to the graph
                'If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                rtnX(i) = freq
                rtnY(i) = ypt
                ' next point...
                i = i + 1
                If i = 1310 Then
                    Do While i < Pts
                        If startf = 0 Then
                            If freq = 0 Then
                                freq = Star
                            Else
                                freq = Star + i * ((Sto - Star) / (Pts - 1))
                            End If
                        Else
                            If freq = 0 Then
                                freq = startf
                            Else
                                freq = startf + i * ((stopf - startf) / (Pts - 1))
                            End If

                        End If
                        ' get the corresponding y points
                        tmp = Mid(Buffer2, (i - 1310) * 40 + 1, 40)
                        ypt = CDbl(Val(tmp))
                        ' add the point to the graph
                        If TraceChecked And Not TweakMode Then t1.AddPoint(TraceID, i, freq, ypt)
                        rtnX(i) = freq
                        rtnY(i) = ypt
                        ' next point...
                        i = i + 1
                    Loop
                    Exit Do
                End If
            Loop
            'If TraceChecked Then t1.AddAllPoints(TraceID, i, rtnX, rtnY)
        ElseIf VNAStr = "N3383A" Then

            x = InStr(ActiveTitle, "RETURN")
            If InStr(ActiveTitle, "RETURN") > 0 Then
                Buffer = DeviceQuery(":CALC1:DATA? FDATA", True)
            Else
                Buffer = DeviceQuery(":CALC1:DATA? FDATA", True)
            End If
            Do While i < Pts
                If startf = 0 Then
                    If freq = 0 Then
                        freq = Star
                    Else
                        freq = Star + i * ((Sto - Star) / (Pts - 1))
                    End If
                Else
                    If freq = 0 Then
                        freq = startf
                    Else
                        freq = startf + i * ((stopf - startf) / (Pts - 1))
                    End If

                End If

                ' get the corresponding y points
                tmp = Mid(Buffer, i * 40 + 1, 40)
                ypt = CDbl(Val(tmp))
                ' add the point to the graph
                ' If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                rtnX(i) = freq
                rtnY(i) = ypt
                ' next point...
                i = i + 1
                If i = 1310 Then
                    Do While i < Pts
                        If startf = 0 Then
                            If freq = 0 Then
                                freq = Star
                            Else
                                freq = Star + i * ((Sto - Star) / (Pts - 1))
                            End If
                        Else
                            If freq = 0 Then
                                freq = startf
                            Else
                                freq = startf + i * ((stopf - startf) / (Pts - 1))
                            End If

                        End If
                        ' get the corresponding y points
                        tmp = Mid(Buffer2, (i - 1310) * 40 + 1, 40)
                        ypt = CDbl(Val(tmp))
                        ' add the point to the graph
                        'If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                        rtnX(i) = freq
                        rtnY(i) = ypt
                        ' next point...
                        i = i + 1
                    Loop
                    Exit Do
                End If
            Loop
            ' If TraceChecked Then t1.AddAllPoints(TraceID, i, rtnX, rtnY)
        Else
            Buffer = DeviceQuery("FORM4;OUTPFORM", True)
            Do While i < Pts
                If startf = 0 Then
                    If freq = 0 Then
                        freq = Star
                    Else
                        freq = Star + i * ((Sto - Star) / (Pts))
                    End If
                Else
                    If freq = 0 Then
                        freq = startf
                    Else
                        freq = startf + i * ((stopf - startf) / (Pts))
                    End If

                End If
                ' get the corresponding y points
                tmp = Mid(Buffer, i * 50 + 1, 50)
                ypt = CDbl(Val(tmp))
                ' add the point to the graph
                ' If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                rtnX(i) = freq
                rtnY(i) = ypt
                ' next point...
                i = i + 1
                If i = 1310 Then
                    Do While i < Pts
                        If startf = 0 Then
                            If freq = 0 Then
                                freq = Star
                            Else
                                freq = Star + i * ((Sto - Star) / (Pts))
                            End If
                        Else
                            If freq = 0 Then
                                freq = startf
                            Else
                                freq = startf + i * ((stopf - startf) / (Pts))
                            End If

                        End If
                        ' get the corresponding y points
                        tmp = Mid(Buffer2, (i - 1310) * 50 + 1, 50)
                        ypt = CDbl(Val(tmp))
                        ' add the point to the graph
                        ' If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                        rtnX(i) = freq
                        rtnY(i) = ypt
                        ' next point...
                        i = i + 1
                    Loop
                    Exit Do
                End If
            Loop
            ' If TraceChecked Then t1.AddAllPoints(TraceID, i, rtnX, rtnY)
        End If
        System.Threading.Thread.Sleep(100)


    End Function
    
    Public Function GetTrace_RL(rtnX() As Double, rtnY() As Double, Optional startf As Double = 0, Optional stopf As Double = 0) As Long
        Dim Buffer As String = ""
        Dim Buffer2 As String = ""
        Dim i As Long
        Dim freq As Double
        Dim tmp As String
        Dim ypt As Double
        Dim t1 As New Trace
        Dim TraceID As Long
        Dim x As Integer
        Array.Clear(rtnY, 0, Pts)
        Array.Clear(rtnX, 0, Pts)

        i = 0
        Pts = 201
        If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(1000)

        If EAveraging Then
            Averages = frmAUTOTEST.AvgS.Text
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite("SENS:AVER OFF")
                ScanGPIB.BusWrite("SENS:AVER ON")
                ScanGPIB.BusWrite("SENS2:AVER OFF")
                ScanGPIB.BusWrite("SENS2:AVER ON")
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("SENS:AVER OFF")
                ScanGPIB.BusWrite("SENS:AVER ON")
                ScanGPIB.BusWrite("SENS2:AVER OFF")
                ScanGPIB.BusWrite("SENS2:AVER ON")
            Else
                ScanGPIB.BusWrite("OPC?;CHAN1")
                ScanGPIB.BusWrite("AVERO OFF")
                ScanGPIB.BusWrite("AVERO ON")
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("AVERO OFF")
                ScanGPIB.BusWrite("AVERO ON")
                ScanGPIB.BusWrite("OPC?;CHAN1")
            End If
            System.Threading.Thread.Sleep(Averages * 500)
        End If
        GetTrace_RL = 0
        'If TraceChecked And Not TweakMode Then
        '    TraceID = t1.OpenTrace
        '    GetTrace = TraceID
        'End If
        If VNAStr = "HP_8720ES" Then Pts = Pts - 1
        ' tell the vna to send the ascii data
        If VNAStr = "AG_E5071B" Then
            Buffer = DeviceQuery(":CALC1:DATA:FDAT?", True)
            Do While i < Pts
                If startf = 0 Then
                    If freq = 0 Then
                        freq = Star
                    Else
                        freq = Star + i * ((Sto - Star) / (Pts - 1))
                    End If
                Else
                    If freq = 0 Then
                        freq = startf
                    Else
                        freq = startf + i * ((stopf - startf) / (Pts - 1))
                    End If

                End If
                ' get the corresponding y points
                tmp = Mid(Buffer, i * 40 + 1, 40)
                ypt = CDbl(Val(tmp))
                ' add the point to the graph
                'If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                rtnX(i) = freq
                rtnY(i) = ypt
                ' next point...
                i = i + 1
                If i = 1310 Then
                    Do While i < Pts
                        If startf = 0 Then
                            If freq = 0 Then
                                freq = Star
                            Else
                                freq = Star + i * ((Sto - Star) / (Pts - 1))
                            End If
                        Else
                            If freq = 0 Then
                                freq = startf
                            Else
                                freq = startf + i * ((stopf - startf) / (Pts - 1))
                            End If

                        End If
                        ' get the corresponding y points
                        tmp = Mid(Buffer2, (i - 1310) * 40 + 1, 40)
                        ypt = CDbl(Val(tmp))
                        ' add the point to the graph
                        If TraceChecked And Not TweakMode Then t1.AddPoint(TraceID, i, freq, ypt)
                        rtnX(i) = freq
                        rtnY(i) = ypt
                        ' next point...
                        i = i + 1
                    Loop
                    Exit Do
                End If
            Loop
            'If TraceChecked Then t1.AddAllPoints(TraceID, i, rtnX, rtnY)
        ElseIf VNAStr = "N3383A" Then

            x = InStr(ActiveTitle, "RETURN")
            If InStr(ActiveTitle, "RETURN") > 0 Then
                Buffer = DeviceQuery(":CALC1:DATA? FDATA", True)
            Else
                Buffer = DeviceQuery(":CALC1:DATA? FDATA", True)
            End If
            Do While i < Pts
                If startf = 0 Then
                    If freq = 0 Then
                        freq = Star
                    Else
                        freq = Star + i * ((Sto - Star) / (Pts - 1))
                    End If
                Else
                    If freq = 0 Then
                        freq = startf
                    Else
                        freq = startf + i * ((stopf - startf) / (Pts - 1))
                    End If

                End If

                ' get the corresponding y points
                tmp = Mid(Buffer, i * 40 + 1, 40)
                ypt = CDbl(Val(tmp))
                ' add the point to the graph
                ' If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                rtnX(i) = freq
                rtnY(i) = ypt
                ' next point...
                i = i + 1
                If i = 1310 Then
                    Do While i < Pts
                        If startf = 0 Then
                            If freq = 0 Then
                                freq = Star
                            Else
                                freq = Star + i * ((Sto - Star) / (Pts - 1))
                            End If
                        Else
                            If freq = 0 Then
                                freq = startf
                            Else
                                freq = startf + i * ((stopf - startf) / (Pts - 1))
                            End If

                        End If
                        ' get the corresponding y points
                        tmp = Mid(Buffer2, (i - 1310) * 40 + 1, 40)
                        ypt = CDbl(Val(tmp))
                        ' add the point to the graph
                        'If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                        rtnX(i) = freq
                        rtnY(i) = ypt
                        ' next point...
                        i = i + 1
                    Loop
                    Exit Do
                End If
            Loop
            ' If TraceChecked Then t1.AddAllPoints(TraceID, i, rtnX, rtnY)
        Else
            Buffer = DeviceQuery("FORM4;OUTPFORM", True)
            Do While i < Pts
                If startf = 0 Then
                    If freq = 0 Then
                        freq = Star
                    Else
                        freq = Star + i * ((Sto - Star) / (Pts))
                    End If
                Else
                    If freq = 0 Then
                        freq = startf
                    Else
                        freq = startf + i * ((stopf - startf) / (Pts))
                    End If

                End If
                ' get the corresponding y points
                tmp = Mid(Buffer, i * 50 + 1, 50)
                ypt = CDbl(Val(tmp))
                ' add the point to the graph
                ' If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                rtnX(i) = freq
                rtnY(i) = ypt
                ' next point...
                i = i + 1
                If i = 1310 Then
                    Do While i < Pts
                        If startf = 0 Then
                            If freq = 0 Then
                                freq = Star
                            Else
                                freq = Star + i * ((Sto - Star) / (Pts))
                            End If
                        Else
                            If freq = 0 Then
                                freq = startf
                            Else
                                freq = startf + i * ((stopf - startf) / (Pts))
                            End If

                        End If
                        ' get the corresponding y points
                        tmp = Mid(Buffer2, (i - 1310) * 50 + 1, 50)
                        ypt = CDbl(Val(tmp))
                        ' add the point to the graph
                        ' If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                        rtnX(i) = freq
                        rtnY(i) = ypt
                        ' next point...
                        i = i + 1
                    Loop
                    Exit Do
                End If
            Loop
            ' If TraceChecked Then t1.AddAllPoints(TraceID, i, rtnX, rtnY)
        End If
        System.Threading.Thread.Sleep(100)


    End Function
    Public Function GetTrace_ext(rtnX() As Double, rtnY() As Double, str As Double, stp As Double) As Long
        Dim Buffer As String = ""
        Dim Buffer2 As String = ""
        Dim i As Long
        Dim freq As Double
        Dim tmp As String
        Dim ypt As Double
        Dim t1 As New Trace
        Dim TraceID As Long
        Dim x As Integer
        Array.Clear(rtnY, 0, Pts)
        Array.Clear(rtnX, 0, Pts)

        i = 0
        Pts = 201
        If VNAStr = "HP_8753C" Then System.Threading.Thread.Sleep(1000)

        If EAveraging Then
            Averages = frmAUTOTEST.AvgS.Text
            If VNAStr = "AG_E5071B" Then
                ScanGPIB.BusWrite("SENS:AVER OFF")
                ScanGPIB.BusWrite("SENS:AVER ON")
                ScanGPIB.BusWrite("SENS2:AVER OFF")
                ScanGPIB.BusWrite("SENS2:AVER ON")
            ElseIf VNAStr = "N3383A" Then
                ScanGPIB.BusWrite("SENS:AVER OFF")
                ScanGPIB.BusWrite("SENS:AVER ON")
                ScanGPIB.BusWrite("SENS2:AVER OFF")
                ScanGPIB.BusWrite("SENS2:AVER ON")
            Else
                ScanGPIB.BusWrite("OPC?;CHAN1")
                ScanGPIB.BusWrite("AVERO OFF")
                ScanGPIB.BusWrite("AVERO ON")
                ScanGPIB.BusWrite("OPC?;CHAN2;")
                ScanGPIB.BusWrite("AVERO OFF")
                ScanGPIB.BusWrite("AVERO ON")
            End If
            System.Threading.Thread.Sleep(Averages * 500)
        End If
        GetTrace_ext = 0
        'If TraceChecked And Not TweakMode Then
        '    TraceID = t1.OpenTrace
        '    GetTrace = TraceID
        'End If
        If VNAStr = "HP_8720ES" Then Pts = Pts - 1
        ' tell the vna to send the ascii data
        If VNAStr = "AG_E5071B" Then
            Buffer = DeviceQuery(":CALC1:DATA:FDAT?", True)
            Do While i < Pts
                freq = Start + i * ((stp - str) / (Pts - 1))

                ' get the corresponding y points
                tmp = Mid(Buffer, i * 40 + 1, 40)
                ypt = CDbl(Val(tmp))
                ' add the point to the graph
                'If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                rtnX(i) = freq
                rtnY(i) = ypt
                ' next point...
                i = i + 1
                If i = 1310 Then
                    Do While i < Pts
                        freq = Start + i * ((stp - str) / (Pts - 1))
                        ' get the corresponding y points
                        tmp = Mid(Buffer2, (i - 1310) * 40 + 1, 40)
                        ypt = CDbl(Val(tmp))
                        ' add the point to the graph
                        If TraceChecked And Not TweakMode Then t1.AddPoint(TraceID, i, freq, ypt)
                        rtnX(i) = freq
                        rtnY(i) = ypt
                        ' next point...
                        i = i + 1
                    Loop
                    Exit Do
                End If
            Loop
            'If TraceChecked Then t1.AddAllPoints(TraceID, i, rtnX, rtnY)
        ElseIf VNAStr = "N3383A" Then

            x = InStr(ActiveTitle, "RETURN")
            If InStr(ActiveTitle, "RETURN") > 0 Then
                Buffer = DeviceQuery(":CALC1:DATA? FDATA", True)
            Else
                Buffer = DeviceQuery(":CALC1:DATA? FDATA", True)
            End If
            Do While i < Pts
                freq = Start + i * ((stp - str) / (Pts - 1))

                ' get the corresponding y points
                tmp = Mid(Buffer, i * 40 + 1, 40)
                ypt = CDbl(Val(tmp))
                ' add the point to the graph
                ' If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                rtnX(i) = freq
                rtnY(i) = ypt
                ' next point...
                i = i + 1
                If i = 1310 Then
                    Do While i < Pts
                        freq = Start + i * ((Sto - Star) / (Pts - 1))
                        ' get the corresponding y points
                        tmp = Mid(Buffer2, (i - 1310) * 40 + 1, 40)
                        ypt = CDbl(Val(tmp))
                        ' add the point to the graph
                        'If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                        rtnX(i) = freq
                        rtnY(i) = ypt
                        ' next point...
                        i = i + 1
                    Loop
                    Exit Do
                End If
            Loop
            ' If TraceChecked Then t1.AddAllPoints(TraceID, i, rtnX, rtnY)
        Else
            Buffer = DeviceQuery("FORM4;OUTPFORM", True)
            Do While i < Pts
                freq = Star + i * ((Sto - Star) / (Pts))

                ' get the corresponding y points
                tmp = Mid(Buffer, i * 50 + 1, 50)
                ypt = CDbl(Val(tmp))
                ' add the point to the graph
                ' If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                rtnX(i) = freq
                rtnY(i) = ypt
                ' next point...
                i = i + 1
                If i = 1310 Then
                    Do While i < Pts
                        freq = Start + i * ((stp - str) / (Pts))
                        ' get the corresponding y points
                        tmp = Mid(Buffer2, (i - 1310) * 50 + 1, 50)
                        ypt = CDbl(Val(tmp))
                        ' add the point to the graph
                        ' If TraceChecked Then t1.AddPoint(TraceID, i, freq, ypt)
                        rtnX(i) = freq
                        rtnY(i) = ypt
                        ' next point...
                        i = i + 1
                    Loop
                    Exit Do
                End If
            Loop
            ' If TraceChecked Then t1.AddAllPoints(TraceID, i, rtnX, rtnY)
        End If
        System.Threading.Thread.Sleep(100)


    End Function

    Public Sub GetTraceMem()
        ' tell the vna to send the ascii data


        If VNAStr = "AG_E5071B" Then
            ScanGPIB.BusWrite(":CALC1:PAR2:SEL")
            gBuffer = DeviceQuery(":CALC1:DATA:SMEM?", True)
        ElseIf VNAStr = "N3383A" Then
            ScanGPIB.BusWrite("CALC1:PAR:SEL 'CH1_S21_1'")
            gBuffer = DeviceQuery(":CALC1:DATA:SMEM?", True)
        Else
            gBuffer = DeviceQuery("OUTPMEMO;", True)
        End If
        System.Threading.Thread.Sleep(100)

    End Sub

    Public Sub SaveCalCoeff(SwPos As Integer)
        Dim Buffer As String
        Dim CoeffCalFile As String
        Dim Title As String


        On Error GoTo Trap
        Title = ActiveTitle
        ActiveTitle = "     SAVING CALIBRATION COEFFICIENTS   "

        'S11 Channel 1
        System.Threading.Thread.Sleep(3000)
        ScanGPIB.BusWrite("OPC?;CHAN1;")
        System.Threading.Thread.Sleep(2000)
        'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
        ' tell the vna to send the ascii data
        ScanGPIB.BusWrite("FORM3;")
        System.Threading.Thread.Sleep(2000)
        Buffer = DeviceQuery("OUTPCALC01;", True)
        CoeffCalFile = CoffCalBasePath & "S11_1_" & SwPos & ".txt"
        LoadCoeffCalToFile(CoeffCalFile, Buffer)

        ScanGPIB.BusWrite("FORM3;")
        System.Threading.Thread.Sleep(2000)
        Buffer = DeviceQuery("OUTPCALC02;", True)
        CoeffCalFile = CoffCalBasePath & "S11_2_" & SwPos & ".txt"
        LoadCoeffCalToFile(CoeffCalFile, Buffer)

        ScanGPIB.BusWrite("FORM3;")
        System.Threading.Thread.Sleep(2000)
        Buffer = DeviceQuery("OUTPCALC03;", True)
        CoeffCalFile = CoffCalBasePath & "S11_1_" & SwPos & ".txt"
        LoadCoeffCalToFile(CoeffCalFile, Buffer)


        'S21 Channel 2
        System.Threading.Thread.Sleep(1000)
        ScanGPIB.BusWrite("OPC?;CHAN2;")
        System.Threading.Thread.Sleep(2000)
        'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
        System.Threading.Thread.Sleep(1000)
        ' tell the vna to send the ascii data
        ScanGPIB.BusWrite("FORM3;")
        System.Threading.Thread.Sleep(2000)
        Buffer = DeviceQuery("OUTPCALC01;", True)
        CoeffCalFile = CoffCalBasePath & "S21_" & SwPos & ".txt"
        LoadCoeffCalToFile(CoeffCalFile, Buffer)
        ActiveTitle = Title
        Exit Sub
Trap:
        ActiveTitle = Title
    End Sub


    Public Sub LoadCalCoeff(SwPos As Integer)
        Dim Buffer As String
        Dim CoeffCalFile As String
        Dim Title As String


        On Error GoTo Trap
        Title = ActiveTitle
        ActiveTitle = "     LOADING CALIBRATION COEFFICIENTS   "

        'S11 Channel 1
        System.Threading.Thread.Sleep(1000)
        ScanGPIB.BusWrite("OPC?;CHAN1;")
        'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
        ' tell the vna to send the ascii data
        CoeffCalFile = CoffCalBasePath & "S11_1_" & SwPos & ".txt"
        Buffer = ReadCoeffCalFromFile(CoeffCalFile)
        ScanGPIB.BusWrite("FORM3;")
        System.Threading.Thread.Sleep(2000)
        ScanGPIB.BusWrite("INPUCALC01, " & Buffer & ";") ' Input CoeffCal data to VNA Memory
        ScanGPIB.BusRead()

        CoeffCalFile = CoffCalBasePath & "S11_2_" & SwPos & ".txt"
        Buffer = ReadCoeffCalFromFile(CoeffCalFile)
        ScanGPIB.BusWrite("OPC?;INPUCALC02," & Buffer & ";") ' Input CoeffCal data to VNA Memory
        ScanGPIB.BusRead()

        CoeffCalFile = CoffCalBasePath & "S11_3_" & SwPos & ".txt"
        Buffer = ReadCoeffCalFromFile(CoeffCalFile)
        ScanGPIB.BusWrite("OPC?;INPUCALC03," & Buffer & ";") ' Input CoeffCal data to VNA Memory
        ScanGPIB.BusRead()

        'S21 Channel 2
        System.Threading.Thread.Sleep(1000)
        ScanGPIB.BusWrite("OPC?;CHAN2;")
        System.Threading.Thread.Sleep(1000)
        'If VNAStr = "HP_8753C" Then ScanGPIB.BusRead()
        System.Threading.Thread.Sleep(1000)
        ' tell the vna to send the ascii data
        CoeffCalFile = CoffCalBasePath & "S21_" & SwPos & ".txt"
        Buffer = ReadCoeffCalFromFile(CoeffCalFile)
        ScanGPIB.BusWrite("OPC?;INPUCALC01," & Buffer & ";") ' Input CoeffCal data to VNA Memory
        ScanGPIB.BusRead()
        ActiveTitle = Title
        Exit Sub
Trap:
        ActiveTitle = Title
    End Sub

    Public Function GetTimeout() As Long
        GetTimeout = 6000
        If VNAStr.Contains("N3383A") Then GetTimeout = 6000
        If VNAStr.Contains("E5071B") Then GetTimeout = 6000
        If VNAStr.Contains("8720ES") Then GetTimeout = 6000
        If VNAStr.Contains("8753E") Then GetTimeout = 6000
        If VNAStr.Contains("8753C") Then GetTimeout = 10000
    End Function


End Module
