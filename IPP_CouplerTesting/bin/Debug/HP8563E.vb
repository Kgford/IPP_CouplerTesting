Option Strict On

Friend Class HP8563E
    Inherits SPA

    Public Shadows myspecan As DirectIO.DirectIO
    Private avg_8563 As Boolean
    Private rb_8563 As Boolean
    Private vb_8563 As Boolean
    Private sw_8563 As Boolean
    Sub New()

    End Sub
    Overrides Function connect(ByVal address As String, ByVal timeout As Integer) As Boolean
        Try
            myspecan = New DirectIO.DirectIO(address, False, 5000)
            myspecan.Timeout = timeout
            myspecan.WriteLine("IP")
            'myspecan.WriteLine("TDF P")
        Catch ex As Exception

        End Try
    End Function

    Overrides Function reset() As Boolean
        myspecan.WriteLine("IP")
    End Function

    Overrides Function setCFreq(ByVal cFreq As Double) As Boolean
        myspecan.WriteLine("CF " & CStr(Math.Round(cFreq, 0)) & "HZ")
    End Function

    Overrides Function GetCFreq() As Long
        myspecan.WriteLine("CF?")
        GetCFreq = CLng(myspecan.Read())
    End Function

    Overrides Function GetStartFreq() As Long
        myspecan.WriteLine("FA?")
        GetStartFreq = CLng(myspecan.Read())
    End Function

    Overrides Function GetStopFreq() As Long
        myspecan.WriteLine("FB?")
        GetStopFreq = CLng(myspecan.Read())
    End Function

    Overrides Function setSweepTime(ByVal mySweep As Double) As Boolean
        myspecan.WriteLine("ST " & CStr(Math.Round(mySweep, 0)) & "MS")
        sw_8563 = False

    End Function

    Overrides Function GetSweepTime() As Double
        myspecan.WriteLine("ST?")
        GetSweepTime = Val(myspecan.Read())
    End Function

    Overrides Function setSweepAuto(ByVal mySWAuto As Integer) As Boolean
        If mySWAuto = 1 Then
            sw_8563 = True
            myspecan.WriteLine("ST AUTO")
        Else
            myspecan.WriteLine("ST MAN")
            sw_8563 = False
        End If

    End Function

    Overrides Function GetSweepAuto() As Boolean
        'myspecan.WriteLine("SWE:TIME:AUTO?")
        'If CInt(myspecan.Read()) = 1 Then
        '    GetSweepAuto = True
        'Else
        GetSweepAuto = sw_8563
        'End If
    End Function

    Overrides Function setSpan(ByVal mySpan As Double) As Boolean
        myspecan.WriteLine("SP " & CStr(Math.Round(mySpan, 0)) & "HZ")

    End Function

    Overrides Function GetSpan() As Long
        myspecan.WriteLine("SP?")
        GetSpan = CLng(myspecan.Read())
    End Function

    Overrides Function setRBW(ByVal myRBW As Long) As Boolean
        rb_8563 = False
        myspecan.WriteLine("RB " & CStr(Math.Round(myRBW, 0)) & "HZ")

    End Function

    Overrides Function GetRBW() As Long
        myspecan.WriteLine("RB?")
        GetRBW = CLng(myspecan.Read())
    End Function

    Overrides Function setRBWAuto(ByVal myRBWAuto As Integer) As Boolean
        If myRBWAuto = 1 Then
            rb_8563 = True
            myspecan.WriteLine("RB AUTO")
        Else
            rb_8563 = False
            myspecan.WriteLine("RB MAN")
        End If

    End Function

    Overrides Function GetRBWAuto() As Boolean
        'myspecan.WriteLine("BWID:AUTO?")
        'If CInt(myspecan.Read()) = 1 Then
        '    GetRBWAuto = True
        'Else
        GetRBWAuto = rb_8563
        'End If
    End Function

    Overrides Function setVBW(ByVal myVBW As Long) As Boolean
        vb_8563 = False
        myspecan.WriteLine("VB " & CStr(Math.Round(myVBW, 0)) & "HZ")

    End Function

    Overrides Function GetVBW() As Long
        myspecan.WriteLine("VB?")
        GetVBW = CLng(myspecan.Read())
    End Function

    Overrides Function setVBWAuto(ByVal myVBWAuto As Integer) As Boolean
        If myVBWAuto = 1 Then
            vb_8563 = True
            myspecan.WriteLine("VB AUTO")
        Else
            vb_8563 = False
            myspecan.WriteLine("VB MAN")
        End If
    End Function

    Overrides Function GetVBWAuto() As Boolean
        'myspecan.WriteLine("BWID:VID:AUTO?")
        'If CInt(myspecan.Read()) = 1 Then
        '    GetVBWAuto = True
        'Else
        GetVBWAuto = vb_8563
        'End If
    End Function

    Overrides Function setAverOn(ByVal myAverOn As Integer) As Boolean
        If myAverOn = 1 Then
            myspecan.WriteLine("VAVG ON")
            avg_8563 = True
        Else
            myspecan.WriteLine("VAVG OFF")
            avg_8563 = False
        End If
    End Function

    Overrides Function GetAverOn() As Boolean
        'myspecan.WriteLine("AVER?")
        'If CInt(myspecan.Read()) = 1 Then
        '    GetAverOn = True
        'Else
        '    GetAverOn = False
        'End If
        GetAverOn = avg_8563
    End Function

    Overrides Function setAverNo(ByVal myAverNo As Integer) As Boolean
        myspecan.WriteLine("VAVG " & CStr(myAverNo))
        avg_8563 = True

    End Function

    Overrides Function GetAverNo() As Integer
        myspecan.WriteLine("VAVG?")
        GetAverNo = CInt(myspecan.Read())
    End Function

    Overrides Function setPdiv(ByVal myPdiv As Double) As Boolean
        myspecan.WriteLine("LG " & CStr(Math.Round(myPdiv, 1)) & "DB")

    End Function

    Overrides Function GetPdiv() As Double
        myspecan.WriteLine("LG?")
        GetPdiv = Val(myspecan.Read())
    End Function

    Overrides Function setRL(ByVal myRL As Double) As Boolean
        myspecan.WriteLine("RL " & CStr(Math.Round(myRL, 3)) & "DBM")

    End Function

    Overrides Function GetRL() As Double
        myspecan.WriteLine("RL?")
        GetRL = Val(myspecan.Read())
    End Function

    Overrides Function setATT(ByVal myATT As Integer) As Integer
        myspecan.WriteLine("AT " & CStr(myATT) & "DB")

    End Function

    Overrides Function GetATT() As Integer
        myspecan.WriteLine("AT?")
        GetATT = CInt(myspecan.Read())
    End Function

    Overrides Function setATTAuto(ByVal myATTAuto As Integer) As Boolean
        If myATTAuto = 1 Then
            myspecan.WriteLine("AT AUTO")
        Else
            myspecan.WriteLine("AT MAN")
        End If
    End Function

    Overrides Function setAvgType(ByVal myAvgType As Integer) As Boolean
        'Select Case myAvgType
        '    Case 0
        '        myspecan.WriteLine("AVER:TYPE LOG")
        '    Case 1
        '        myspecan.WriteLine("AVER:TYPE RMS")
        '    Case 2
        '        myspecan.WriteLine("AVER:TYPE SCAL")
        'End Select

    End Function

    Overrides Function GetAvgType() As Integer
        GetAvgType = 0
    End Function

    Overrides Function GetATTAuto() As Boolean
        'myspecan.WriteLine("POW:ATT:AUTO?")
        'If CInt(myspecan.Read()) = 1 Then
        '    GetATTAuto = True
        'Else
        '    GetATTAuto = False
        'End If
    End Function

    Overrides Function GetTraceMode(ByVal myTrace As Integer) As Integer
        If myTrace = 1 Then
            Select Case TRAmode
                Case "CLRW"
                    GetTraceMode = 0
                Case "MXMH"
                    GetTraceMode = 1
                Case "MINH"
                    GetTraceMode = 2
                Case "VIEW"
                    GetTraceMode = 3
                Case "BLANK"
                    GetTraceMode = 4
            End Select
        Else
            Select Case TRBmode
                Case "CLRW"
                    GetTraceMode = 0
                Case "MXMH"
                    GetTraceMode = 1
                Case "MINH"
                    GetTraceMode = 2
                Case "VIEW"
                    GetTraceMode = 3
                Case "BLANK"
                    GetTraceMode = 4
            End Select
        End If
 

    End Function

    Overrides Function setTraceMode(ByVal trace As Integer, ByVal mode As Integer) As Boolean
        Select Case trace
            Case 1
                Select Case mode
                    Case 0
                        myspecan.WriteLine("CLRW TRA")
                        TRAmode = "CLRW"
                    Case 1
                        myspecan.WriteLine("MXMH TRA")
                        TRAmode = "MXMH"
                    Case 2
                        myspecan.WriteLine("MINH TRA")
                        TRAmode = "MINH"
                    Case 3
                        myspecan.WriteLine("VIEW TRA")
                        TRAmode = "VIEW"
                    Case 4
                        myspecan.WriteLine("BLANK TRA")
                        TRAmode = "BLANK"
                End Select
            Case 2
                Select Case mode
                    Case 0
                        myspecan.WriteLine("CLRW TRB")
                        TRBmode = "CLRW"
                    Case 1
                        myspecan.WriteLine("MXMH TRB")
                        TRBmode = "MXMH"
                    Case 2
                        myspecan.WriteLine("MINH TRB")
                        TRBmode = "MINH"
                    Case 3
                        myspecan.WriteLine("VIEW TRB")
                        TRBmode = "VIEW"
                    Case 4
                        myspecan.WriteLine("BLANK TRB")
                        TRBmode = "BLANK"
                End Select
        End Select
  
    End Function


    Overrides Function setSweepOff(ByVal avg As Boolean) As Boolean
        myspecan.WriteLine("SNGLS")

    End Function

    Overrides Function setSweepOn() As Boolean
        myspecan.WriteLine("CONTS")
    End Function

    Overrides Function setTakeSweep(ByVal avg As Boolean) As Boolean
        Dim scheck As Integer
        myspecan.WriteLine("TS")
        scheck = CInt(myspecan.ReadStatusByte)
        'myspecan.WriteLine("RQS 16")
    End Function

    Overrides Function setBlank() As Boolean
        myspecan.WriteLine("BLANK TRA")
    End Function

    Overrides Function setClearWrite() As Boolean
        myspecan.WriteLine("TRAC:MODE WRIT")
    End Function

    Overrides Function getOperationComplete() As Boolean
        Dim sComplete As String
        getOperationComplete = False
        myspecan.WriteLine("DONE?")
        sComplete = myspecan.Read()
        If CInt(sComplete) = 1 Then getOperationComplete = True
    End Function

    Overrides Function getStatus(ByVal avg As Boolean) As Boolean
        Dim sCheck As Integer
        Static myDone As Boolean
        If myDone = True Then
            myDone = False
            getStatus = True
            Exit Function
        End If
        getStatus = False

        'myspecan.WriteLine("*OPC")
        'myspecan.WriteLine("*STB?")
        sCheck = CInt(myspecan.ReadStatusByte)
        If (sCheck And 16) = 16 Then
            getStatus = True
            myDone = True
        End If
    End Function
    'Overrides Function GetTrace(ByVal ref As Double, ByVal scale As Double) As String
    '    Dim myval As String
    '    myspecan.WriteLine("TRAC? TRACE1")
    '    myval = myspecan.Read()
    '    GetTrace = myval
    'End Function
    Overrides Function GetTrace(ByVal trace As Integer, ByVal ref As Double, ByVal scale As Double) As Double()
        Dim myval As Byte()
        Dim temp As Double()
        Dim y As Integer
        Dim x As Integer
        Dim a As Integer
        If trace = 1 Then
            myspecan.WriteLine("TDF B;TRA?")
        ElseIf trace = 2 Then
            myspecan.WriteLine("TDF B;TRB?")

        End If
        myval = myspecan.UnbufferedRead(1202)
        x = 0
        y = 0
        For a = 0 To 600
            ReDim Preserve temp(y)
            temp(y) = ref + scale * (((myval(x) * 256) + myval(x + 1)) / 60 - 10)
            x = x + 2
            y = y + 1
        Next a
        GetTrace = temp
    End Function

    Overrides Function setLocal() As Boolean
        myspecan.GPIBLocal()
    End Function

    Overrides Function setDetector(ByVal trace As Integer, ByVal det As Integer) As Boolean

    End Function


    Overrides Function getDetector(ByVal trace As Integer) As Integer

    End Function


End Class
