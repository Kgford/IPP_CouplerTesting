Option Explicit On
Imports System.Xml
Imports System.IO
Imports System.Net
Imports System.Text




Module globals

    Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long

    Public Daqboard As New MccDaq.MccBoard()
    Public ulstat As MccDaq.ErrorInfo
    Public Status As Boolean
    Public RobotStatus As Boolean = False
    Public MCCLoaded As Boolean = True
    Public Robot As Boolean = True
    Public RobotError As Boolean = False
    Public RobotMoving As Boolean = False
    Public ReadySignal As Boolean = False
    Public TestRunning As Boolean = False
    Public TestComplete As Boolean = False
    Public TestPassFail As Boolean = False
    Public TestRetest As Integer = 0
    Public Retest As Boolean = False
    Public NewTray As String = "OK"
    Public PassTray As String = "OK"
    Public FailTray As String = "OK"
    Public OperatorLog As Boolean = False
    Public FirstComplete As Boolean = True
    Public SupervisorPassword As String
    Public GlobalFail As Integer
    Public Test1Failing As Boolean = False
    Public Test2Failing As Boolean = False
    Public TEST3Failing As Boolean = False
    Public TEST4Failing As Boolean = False
    Public TEST5Failing As Boolean = False
    Public GlobalFailing As Boolean = False
    Public GlobalFail_bypass As Boolean = False
    Public Test1Fail_bypass As Boolean = False
    Public Test2Fail_bypass As Boolean = False
    Public TEST3Fail_bypass As Boolean = False
    Public TEST4Fail_bypass As Boolean = False
    Public TEST5Fail_bypass As Boolean = False
    Public Retest_bypass As Boolean = False
    Public M_bypass As Integer = 0
    Public Master_bypass As Boolean = False
    Public Percent_bypass As Boolean = False
    Public ArtworkRevision As String = "N/A"
    Public jobSpec As String
    Public PartSpec As String
    Public SwitchModel As String
    Public SwitchPorts As Integer


    Public OverrideLogForm As Boolean
    Public LogTextObject As Object
    Public XLocation
    Public YLocation
    Public XSize As Integer
    Public YSize As Integer

    Public retMax As Double
    Public retMin As Double
    Public retMaxX As Double
    Public retMinX As Double

    'Public NumPortClients(10) As Long
    'Public Comms(10) As MSComm
    Public Debug As Boolean = False
    Public Connected As Boolean = False
    Public ActiveTitle As String
    Public LastMessage As String
    Public gBuffer As String
    Public SpecIL As Double
    Public SpecILL As Double
    Public SpecILH As Double
    Public IL_TF As Boolean = False
    Public SpecIL_exp As Double
    Public SpecIL_start1 As Double
    Public SpecIL_stop1 As Double
    Public SpecIL_start2 As Double
    Public SpecIL_stop2 As Double
    Public SpecPB As Double
    Public SpecRL As Double
    Public SpecISO As Double
    Public SpecISOL As Double
    Public SpecISOH As Double
    Public ISO_TF As Boolean = False
    Public SpecCuttoffFreq As Double
    Public SpecCOUP As Double
    Public SpecCOUPPM As Double
    Public SpecDIRECT As Double
    Public SpecCOUPFLAT As Double
    Public Offset1 As Double
    Public Offset2 As Double
    Public Offset3 As Double
    Public Offset4 As Double
    Public Offset5 As Double
    Public Test1 As Integer
    Public Test2 As Integer
    Public Test3 As Integer
    Public Test4 As Integer
    Public Test5 As Integer
    Public SpecAB As Double
    Public SpecAB_TF As Boolean = False
    Public SpecAB_exp As Double
    Public SpecAB_start1 As Double
    Public SpecAB_stop1 As Double
    Public SpecAB_start2 As Double
    Public SpecAB_stop2 As Double
    Public SpecPorts As Integer
    Public SMD As Boolean
    Public YData As Double
    Public SerialNumber As String
    Public Job As String
    Public Test As String
    Public PPH As Double = 115
    Public Quantity As Double = 215
    Public TimeStart As DateTime
    Public LastTest As Long
    Public ExpectedUUTCount As Integer
    Public ExpectedUUTTime As Long = Nothing
    Public ExpectedProgress2 As Integer
    Public TotalTime As TimeSpan
    Public QuantityTime As TimeSpan
    Public PPHTime As TimeSpan
    Public TestPaused As Boolean
    Public TimePerUUT_Pause As Integer = 5 'Mulatiplier when to pause for time per uut 
    Public TimePerUUT As TimeSpan
    Public Timernow As Long
    Public TimeComplete As TimeSpan
    Public Pausenow As Long
    Public TPP As Double

    Public NetworkAccess As Boolean = False
    Public NoInit As Boolean = False

    Public SpecStartFreq As Double
    Public SpecStopFreq As Double
    Public Part As String
    Public VNAStr As String

    Public TraceChecked As Boolean = False
    Public TraceCheckedChanged As Boolean = False
    Public SwitchedChecked As Boolean = False
    Public MutiCalChecked As Boolean = False
    Public DebugChecked As Boolean = False
    Public PassChecked As Boolean = False
    Public FailChecked As Boolean = False
    Public DBDataChecked As Boolean = False
    Public EAveraging As Boolean = False
    Public Averages As Integer
    Public RetrnVal As Double
    Public RetrnVal1 As Double
    Public IL As Double
    Public RL As Double
    Public AB As Double
    Public AB1 As Double
    Public AB2 As Double
    Public AB1Pass As String
    Public AB2Pass As String
    Public DIR As Double
    Public PB As Double
    Public PB1 As Double
    Public PB2 As Double
    Public CF As Double
    Public ISo As Double
    Public ISoL As Double
    Public ISoH As Double
    Public ISoLPass As String
    Public ISoHPass As String
    Public COuP As Double

    Public DebugLevel As Integer
    Public gCmdPrompt As Boolean
    Public DevCount As Long
    ''Public Devices() As TestEquipDev
    Public PropertyNumbers() As String
    Public gTurnSWPRFOff As Boolean

    Public AutoCreateDevice As Boolean
    Public AutoCreateDeviceSet As Boolean
    Public gConfigName As String
    Public bCheckSafeCommand As Boolean

    'Trace
    Public UUTNum As Integer
    Public FirstPart As Boolean = True
    Public UUTReset As Boolean = False
    Public UUTNum_Reset As Integer
    Public TestID As String
    Public DataID As String
    Public XArray(1001) As Double 'Frequency Data Array 
    Public YArray(1001) As Double 'Amplitude Data Array
    Public YArray1(1001) As Double 'Amplitude Data Array
    Public YArray2(1001) As Double 'Amplitude Data Array
    Public IL_XArray(5, 201) As Double 'Frequency Data Array 
    Public IL1_YArray(5, 201) As Double 'Amplitude Data Array
    Public IL2_YArray(5, 201) As Double 'Amplitude Data Array
    Public RL_XArray(5, 201) As Double 'Frequency Data Array 
    Public RL_YArray(5, 201) As Double 'Amplitude Data Array
    Public DIR_XArray(5, 201) As Double 'Frequency Data Array 
    Public DIR_YArray(5, 201) As Double 'Amplitude Data Array
    Public COUP_XArray(5, 201) As Double 'Frequency Data Array 
    Public COUP1_YArray(5, 201) As Double 'Amplitude Data Array
    Public COUP2_YArray(5, 201) As Double 'Amplitude Data Array
    Public ISO_XArray(5, 201) As Double 'Frequency Data Array 
    Public ISO_YArray(5, 201) As Double 'Amplitude Data Array
    Public AB_XArray(5, 201) As Double 'Frequency Data Array 
    Public AB1_YArray(5, 201) As Double 'Amplitude Data Array
    Public AB2_YArray(5, 201) As Double 'Amplitude Data Array
    Public PB_XArray(5, 201) As Double 'Frequency Data Array 
    Public PB1_YArray(5, 201) As Double 'Amplitude Data Array
    Public PB2_YArray(5, 201) As Double 'Amplitude Data Arra

    Public Title As String


    Public Version As String
    Public xTitle As String
    Public yTitle As String
    Public Notes As String
    Public SCALESIZE As String = 5
    Public ProgTitle As String
    Public ProgVer As String
    Public TraceID As String
    Public ActiveDate As String
    Public ActiveData As Double


    Public Active_InstID As String
    Public Temperature As String
    Public InstrumentInfo As String
    Public CalDue As String
    Public CalDate As String
    Public SpecType As String
    Public SpecID As Long

    Public ModelNumber As String
    Public JobNumber As String
    Public WorkStation As String
    Public SavedWorkStation(2) As String
    Public ReportStatus(2) As String
    Public User As String
    Public SavedUser(2) As String
    Public SavedComplete(2) As String
    Public SavedTotal(2) As String
    Public SavedJob(2) As String
    Public resumeTest As Boolean
    Public stopTest As Boolean = False
    Public TweakMode As Boolean = False
    Public TempUUTNum As Integer
    Public BypassUnchecked As Boolean = False
    Public ReportJob As String = "No Job"


    ' Database Location

    Public Const NetworkDataBasePath = "\\ippdc\Test Automation\Automation Programs\ATE Database" ' IPP Network Site
    'Public Const NetworkDataBasePath = "A:\Automation Programs\ATE Database\" ' IPP Network Site
    'Public Const NetworkDataBasePath = "C:\ATE Database\" ' ATS Network Site

    Public Const ExcelTemplatePath = "\\ippdc\Test Automation\Excel Templates\" ' IPP Network Site
    'Public Const ExcelTemplatePath = "C:\Excel Templates\"  ' ATS Network Site

    Public Const TestDataPath = "\\ippdc\Test Automation\Test Data" ' IPP Network Site
    'Public Const TestDataPath = "C:\Test Data\"  ' ATS Network Site

    ' \\ippfs\Test Automation\ATE_Publish\ new network location
    ' Public Const LocalDataBasePath = "\\ippfs\Test Automation\ATE Database\"  'IPP Local Site
    Public Const CoffCalBasePath = "\\ippdc\Test Automation\ATE Data\CalCoeff\"

    Public Const LocalDataBasePath = "C:\ATE Database\"  'IPP Local Site



    Public GPIBCount As Integer

    Public Star As Double
    Public Sto As Double
    Public Pts As Integer
    Public Start As Double

    Public CurrS1 As String
    Public CurrS2 As String
    Public CurrS3 As String
    Public CurrS4 As String
    Public SQLAccess As Boolean = False
    Public SQLVerified As Boolean = False
    Public CalCoeff1 As String = ""
    Public CalCoeff2 As String = ""
    Public CalCoeff3 As String = ""



    'Public Function ComplexDivide(C1 As Complex, C2 As Complex) As Complex

    '    Dim c As Complex
    '    Dim sq As Double

    '    sq = C2.Real ^ 2 + C2.Imag ^ 2

    '    c.Real = (C1.Real * C2.Real + C1.Imag * C2.Imag) / sq
    '    c.Imag = (C1.Imag * C2.Real - C1.Real * C2.Imag) / sq

    '    ComplexDivide = c

    'End Function



    'Public Function GetDefaultTerms() As terms
    '
    '    Dim T As terms ' iotech terminator structure
    '
    '    T.nChar = 2
    '    T.term1 = 13
    '    T.term2 = 10
    '    T.eoi = 1 ' True
    '
    '    GetDefaultTerms = T
    '
    'End Function


    ' forward the error to the caller
    Public Sub ErrReRaise()
        Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End Sub

    Public Function ListToString(ByVal lst As List(Of String)) As String
        Dim reStr As String = String.Empty
        For i As Integer = 0 To lst.Count - 2
            reStr &= lst.Item(i) & ","
        Next
        reStr &= lst.Item(lst.Count - 1)
        ListToString = reStr
    End Function
    Public Function TruncateDecimal(value As Decimal, precision As Integer) As Decimal
        Dim stepper As Decimal = Math.Pow(10, precision)
        Dim tmp As Decimal = Math.Truncate(stepper * value)
        Return tmp / stepper
    End Function

    Public Function StringToList(ByVal str As String) As List(Of String)
        Dim reLst As New List(Of String)
        reLst = str.Split(",").ToArray.ToList
        StringToList = reLst
    End Function
    Public Function Settled(DataPts() As Double, _
            Tolerance As Double, NumPoints As Long) As Boolean

        ' This routine tests array elements 0 to count - 1

        Dim i As Long
        Dim dMax As Double
        Dim dMin As Double

        If NumPoints < 1 Then
            Settled = False
            Exit Function
        End If

        ' initialize max and min to first point
        dMax = DataPts(0)
        dMin = DataPts(0)

        ' Now look trough rest of array for max and min
        For i = 1 To NumPoints - 1
            If DataPts(i) > dMax Then
                dMax = DataPts(i)
            End If
            If DataPts(i) < dMin Then
                dMin = DataPts(i)
            End If
        Next i

        ' tell the use if the data has settled
        Settled = Tolerance >= Math.Abs(dMax - dMin)

    End Function
    Public Function TrimY(YArr As Double(), Optional offset As Double = 0) As Double()
        Dim TempArray(200) As Double
        If YArr.Count = 201 Then
            TrimY = YArr
        Else
            For x = 0 To 200
                ReDim Preserve TempArray(x)
                TempArray(x) = YArr(x) + offset
            Next
            TrimY = TempArray
        End If
    End Function
    Public Function TrimX(XArr As Double()) As Double()
        Dim TempArray(200) As Double
        If XArr.Count = 201 Then
            TrimX = XArr
        Else
            For x = 0 To 200
                TempArray(x) = XArr(x)
            Next
            TrimX = TempArray
        End If
    End Function
    Public Function SplitTraceY(xArr As Double(), YArr As Double(), str As Double, stp As Double, Optional offset As Double = 0) As Double()
        Dim TempxArray(200) As Double
        Dim TempyArray(200) As Double
        Dim count As Integer = 0
        Dim counted As Boolean = False
        For x = 0 To 200
            If str > xArr(0) Then
                If x = 0 Then
                    If xArr(x) >= str And xArr(x + 1) <= stp Then
                        If Not counted Then count += 1
                        counted = True
                        ReDim Preserve TempyArray(x - count)
                        TempyArray(x - count) = YArr(x)
                    End If
                ElseIf x = 200 Then
                    If xArr(x - 1) >= str And xArr(x) <= stp And Not xArr(x) <= 0 Then
                        If Not counted Then count += 1
                        counted = True
                        ReDim Preserve TempyArray(x - count)
                        TempyArray(x - count) = YArr(x)
                    End If
                Else
                    If xArr(x) >= str And xArr(x + 1) <= stp Then
                        If Not counted Then count += 1
                        counted = True
                        ReDim Preserve TempyArray(x - count)
                        TempyArray(x - count) = YArr(x)
                    Else
                        count += 1
                    End If
                End If
            Else
                If x = 0 Then
                    If xArr(x) >= str And xArr(x) <= stp Then
                        ReDim Preserve TempyArray(x)
                        TempyArray(x) = YArr(x)
                    End If
                ElseIf x = 200 Then
                    If xArr(x - 1) >= str And xArr(x) <= stp And Not xArr(x) <= 0 Then
                        ReDim Preserve TempyArray(x)
                        TempyArray(x) = YArr(x)
                    End If
                Else
                    If xArr(x - 1) >= str And xArr(x - 1) <= stp Then
                        ReDim Preserve TempyArray(x)
                        TempyArray(x) = YArr(x)
                    End If
                End If

            End If
        Next
        SplitTraceY = TempyArray
    End Function
    Public Function SplitTraceX(xArr As Double(), YArr As Double(), str As Double, stp As Double) As Double()
        Dim TempxArray(200) As Double
        Dim TempyArray(200) As Double
        Dim count As Integer = 0
        Dim count1 As Integer = 0
        Dim counted As Boolean = False
        For x = 0 To 200
            If str > xArr(0) Then
                If x = 0 Then
                    If xArr(x) >= str And xArr(x + 1) <= stp Then
                        If Not counted Then count += 1
                        counted = True
                        ReDim Preserve TempxArray(x - count)
                        TempxArray(x - count) = xArr(x)
                    End If
                ElseIf x = 200 Then
                    If xArr(x - 1) >= str And xArr(x) <= stp And Not xArr(x) <= 0 Then
                        If Not counted Then count += 1
                        counted = True
                        ReDim Preserve TempxArray(x - count)
                        TempxArray(x - count) = xArr(x)
                    End If
                ElseIf xArr(x) >= str And xArr(x + 1) <= stp Then
                    If Not counted Then count += 1
                    counted = True
                    ReDim Preserve TempxArray(x - count)
                    TempxArray(x - count) = xArr(x)
                Else
                    count += 1
                End If
            Else
                If x = 0 Then
                    If xArr(x) >= str And xArr(x) <= stp Then
                        ReDim Preserve TempxArray(x)
                        TempxArray(x) = xArr(x)
                    End If
                ElseIf x = 200 Then
                    If xArr(x - 1) >= str And xArr(x) <= stp And Not xArr(x) <= 0 Then
                        ReDim Preserve TempxArray(x)
                        TempxArray(x) = xArr(x)
                    End If
                Else
                    If xArr(x - 1) >= str And xArr(x - 1) <= stp Then
                        ReDim Preserve TempxArray(x)
                        TempxArray(x) = xArr(x)
                    End If
                End If

            End If

        Next
        SplitTraceX = TempxArray
    End Function

    Public Function Average(DataPts() As Double, numPts As Long) As Double

        ' This routine averages array elements 0 to count - 1

        Dim i As Long
        Dim Total As Double

        ' calculate total
        For i = 0 To numPts - 1
            Total = Total + DataPts(i)
        Next i

        ' return the average
        Average = Total / numPts

    End Function

    Public Function Max(DataPts() As Double) As Double
        Dim intMax, i As Integer

        For i = 0 To UBound(DataPts) - 1
            If DataPts(i) = 0 Then Exit For
            If DataPts(i) > intMax Then
                intMax = DataPts(i)
            End If
        Next
        Max = intMax

    End Function
    Public Function MaxNoZero(DataPts() As Double) As Double
        Dim i As Integer
        Dim intMax As Double

        For i = 0 To UBound(DataPts) - 1
            If DataPts(i) = 0 Then GoTo DontUse
            If i = 0 Then intMax = DataPts(i)
            If DataPts(i) > intMax Then
                intMax = DataPts(i)
            End If
DontUse:
        Next
        MaxNoZero = intMax

    End Function
    Public Function MinNoZero(DataPts() As Double) As Double
        Dim i As Integer
        Dim intMin As Double

        For i = 0 To UBound(DataPts) - 1
            If DataPts(i) = 0 Then GoTo DontUse
            If i = 0 Then intMin = DataPts(i)
            If DataPts(i) < intMin Then
                intMin = DataPts(i)
            End If
DontUse:
        Next
        MinNoZero = intMin

    End Function


    Public Function MaxX(DataPts() As Double) As Integer
        Dim intMax, i As Integer
        MaxX = 0
        For i = 0 To UBound(DataPts) - 1
            If i = 0 Then intMax = DataPts(i)
            If DataPts(i) > intMax Then
                intMax = DataPts(i)
                MaxX = i
            End If
        Next

    End Function


    Public Function Min(DataPts() As Double) As Double
        Dim intMin, i As Integer
        For i = 0 To UBound(DataPts) - 1
            If i = 0 Then intMin = DataPts(i)
            If DataPts(i) < intMin Then
                intMin = DataPts(i)
            End If
        Next

        Min = intMin
    End Function

    Public Function MinX(DataPts() As Double) As Integer
        Dim intMin, i As Integer
        MinX = 0
        For i = 0 To UBound(DataPts) - 1
            If i = 0 Then intMin = DataPts(i)
            If DataPts(i) < intMin Then
                intMin = DataPts(i)
                MinX = i
            End If
        Next
    End Function

    Public Function Log10(Value As Double) As Double
        On Error GoTo Trap
        Log10 = Math.Log(Value) / Math.Log(10)
        Exit Function
Trap:
    End Function

    Public Function MagnitudeFromRealImag(RealPart As Single, ImagPart As Single) As Single

        'Dim Angle As Single

        ' calculate the magnitude from the real and imaginary pairs
        MagnitudeFromRealImag = (ImagPart ^ 2 + RealPart ^ 2) ^ (0.5)
        'angle = arctan(imag/real)

    End Function
    Public Function VSWRtoRL(VSWR As Double) As Double

        VSWRtoRL = 20 * Log10((VSWR - 1) / (VSWR + 1))

    End Function
    Public Function NoRepeats(data() As String, RecentData As String) As Boolean
        Dim i As Integer

        NoRepeats = True
        For i = 0 To UBound(data)
            If InStr(data(i), RecentData) <> 0 Then
                NoRepeats = False
                Exit Function
            End If
        Next

    End Function

    Public Function StringifyTrace(trace() As Double) As String
        StringifyTrace = ""
        For i = 0 To UBound(trace) - 1
            If StringifyTrace = "" Then
                StringifyTrace = Str(trace(i))
            Else
                StringifyTrace = StringifyTrace + " " + Str(trace(i))
            End If
        Next
    End Function
    Public Function ExpandTrace(trace As String) As Double()
        Dim split_trace() = Split(trace, " ")
        Dim Thistrace(100) As Double
        For i = 0 To UBound(split_trace) - 1
            ReDim Preserve XArray(i)
            Thistrace(i) = CDbl(split_trace(i))
        Next
        ExpandTrace = Thistrace
    End Function




    Public Function AccessDatabaseFolder(Table As String) As String

        AccessDatabaseFolder = NetworkDataBasePath
        If NetworkAccess Then
            AccessDatabaseFolder = NetworkDataBasePath
            Select Case Table
                Case "LocalSpecs"
                    AccessDatabaseFolder = LocalDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & Table & ".mdb "
                Case "NetworkSpecs"
                    AccessDatabaseFolder = NetworkDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = NetworkDataBasePath & Table & ".mdb "
                Case "LocalData"
                    AccessDatabaseFolder = LocalDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & Table & ".mdb "
                Case "NetworkData"
                    AccessDatabaseFolder = NetworkDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = NetworkDataBasePath & Table & ".mdb "
                Case "LocalTraceData"
                    AccessDatabaseFolder = LocalDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & Table & ".mdb "
                Case "NetworkTraceData"
                    AccessDatabaseFolder = NetworkDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = NetworkDataBasePath & Table & ".mdb "
                Case "Devices"
                    AccessDatabaseFolder = NetworkDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = NetworkDataBasePath & Table & ".mdb "
                Case "Specifications"
                    AccessDatabaseFolder = NetworkDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = NetworkDataBasePath & Table & ".mdb "
                Case "Effeciency"
                    AccessDatabaseFolder = NetworkDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = NetworkDataBasePath & Table & ".mdb "
            End Select

        Else
            Select Case Table
                Case "LocalSpecs"
                    AccessDatabaseFolder = LocalDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & Table & ".mdb "
                Case "NetworkSpecs"
                    AccessDatabaseFolder = LocalDataBasePath & "LocalSpecs.accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & "LocalSpecs.mdb"
                Case "LocalData"
                    AccessDatabaseFolder = LocalDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & Table & ".mdb "
                Case "NetworkData"
                    AccessDatabaseFolder = LocalDataBasePath & "LocalData.accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & "LocalData.mdb"
                Case "LocalTraceData"
                    AccessDatabaseFolder = LocalDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & Table & ".mdb "
                Case "NetworkTraceData"
                    AccessDatabaseFolder = LocalDataBasePath & "LocalTraceData.accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & "NetworkTraceData.mdb"
                Case "Devices"
                    AccessDatabaseFolder = LocalDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & Table & ".mdb "
                Case "Specifications"
                    AccessDatabaseFolder = LocalDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & Table & ".mdb "
                Case "Effeciency"
                    AccessDatabaseFolder = LocalDataBasePath & Table & ".accdb"
                    If Not FileExists(AccessDatabaseFolder) Then AccessDatabaseFolder = LocalDataBasePath & Table & ".mdb "


            End Select

        End If

    End Function

    Public Function GetVersion() As String
        Try
            GetVersion = ""
            Dim reader As XmlTextReader = New XmlTextReader("IPP_CouplerTesting.exe.manifest")
            Do While (reader.Read())
                Select Case reader.NodeType
                    Case XmlNodeType.Element 'Display beginning of element.
                        If reader.HasAttributes Then 'If attributes exist
                            While reader.MoveToNextAttribute()
                                If reader.Name = "version" Then
                                    Return reader.Value
                                End If
                            End While
                        End If
                End Select
            Loop
        Catch ex As Exception
            GetVersion = System.Windows.Forms.Application.ProductVersion.ToString
        End Try
    End Function

    Public Function GetComputerName() As String
        Try
            GetComputerName = Environment.UserName
        Catch ex As Exception
            GetComputerName = "Automated Test Solutions"
        End Try
    End Function
    Public Function CheckNetworkFolder() As Boolean
        CheckNetworkFolder = System.IO.Directory.Exists(NetworkDataBasePath)
    End Function
    Public Function FileExists(Path As String) As Boolean
        FileExists = System.IO.File.Exists(Path)
    End Function

    Public Sub LoadCoeffCalToFile(FileName As String, Buffer As String)
        Dim fso, f
        fso = CreateObject("Scripting.FileSystemObject")
        f = fso.OpenTextFile(FileName, 2)
        f.Write(Buffer)
        f.Close()

    End Sub

    Public Function ReadCoeffCalFromFile(FileName As String) As String
        Dim objFile, objText, Text

        objFile = CreateObject("Scripting.FileSystemObject")
        objText = objFile.OpenTextFile(FileName)

        Text = objText.ReadAll
        objText.Close()

        objText = Nothing
        objFile = Nothing

        ReadCoeffCalFromFile = Text
    End Function

    Public Sub Delay(DelayTime As Integer)
        System.Threading.Thread.Sleep(DelayTime)
    End Sub

    Public Sub api_requests()
        Try
            Dim request As WebRequest =
              WebRequest.Create("http://inn-autocon:8888/report_queue/")
            ' If required by the server, set the credentials.
            request.Credentials = CredentialCache.DefaultCredentials
            ' Get the response.
            Dim response As WebResponse = request.GetResponse()
            ' Display the status.
            Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
            ' Get the stream containing content returned by the server.
            Dim dataStream As Stream = response.GetResponseStream()
            ' Open the stream using a StreamReader for easy access.
            Dim reader As New StreamReader(dataStream)
            ' Read the content.
            Dim responseFromServer As String = reader.ReadToEnd()
            ' Display the content.
            Console.WriteLine(responseFromServer)
            ' Clean up the streams and the response.
            reader.Close()
            response.Close()
        Catch ex As Exception

        End Try
    End Sub
    Public Sub post_api()
        Dim url As String = "http://inn-autocon:8888/report_queue/"
        Dim uri As New Uri(url)
        Dim jsonSring As String = json_serializer()
        Dim data = Encoding.UTF8.GetBytes(jsonSring)
        Dim result_post = SendRequest(uri, data, "application/json", "POST")
    End Sub
    Public Function SendRequest(uri As Uri, jsonDataBytes As Byte(), contentType As String, method As String) As String
        Dim response As String
        Dim request As WebRequest
        Try
            request = WebRequest.Create(uri)
            request.ContentLength = jsonDataBytes.Length
            request.ContentType = contentType
            request.Method = method

            Using requestStream = request.GetRequestStream
                requestStream.Write(jsonDataBytes, 0, jsonDataBytes.Length)
                requestStream.Close()

                Using responseStream = request.GetResponse.GetResponseStream
                    Using reader As New StreamReader(responseStream)
                        response = reader.ReadToEnd()
                    End Using
                End Using
            End Using
            Return response
        Catch ex As Exception

        End Try

    End Function

    Private Function encode(ByVal str As String) As Byte()
        'supply True as the construction parameter to indicate
        'that you wanted the class to emit BOM (Byte Order Mark)
        'NOTE: this BOM value is the indicator of a UTF-8 string
        Dim utf8Encoding As New System.Text.UTF8Encoding(True)
        Dim encodedString() As Byte

        encodedString = utf8Encoding.GetBytes(str)

        Return encodedString
    End Function

    Private Function json_serializer() As String
        Dim today As DateTime = DateTime.Now
        json_serializer = " 'results': [{'reportname': 'This is a test','reporttype': " & SpecType & ",'reportstatus': 0,'jobnumber':" & JobNumber & ",'workstation': " & WorkStation & ",'partnumber': " & Part & ",'operator': " & User & ",'activedate': " & today & "}]"
    End Function


End Module
