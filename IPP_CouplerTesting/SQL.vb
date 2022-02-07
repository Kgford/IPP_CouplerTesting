

Imports System.Data
Imports System.Data.SqlClient


Module SQL
    Public reader As SqlDataReader
    Public Localreader As SqlDataReader
    Public IDArray(30) As String
    Private MissingID_Array(30) As String
    Private MissingINST_Array(30) As String
    Private installedstr As String
    Private Missingstr As String
    Private MissingStr_Array(30) As String
    Public LibName As String

    Public SQLConnStr As String = "Data Source=INN-SQLEXPRESS\SQLEXPRESS;Initial Catalog=ATE;User Id='developer';Password='secure';" 'Development Server

    Public Function CheckDatabaseExists(ByVal server As String, ByVal database As String) As Boolean
        Try

            Dim connString As String = ("Data Source=" _
                        & (server & ";Initial Catalog=master;Integrated Security=True;"))
            Dim cmdText As String = ("select * from master.dbo.sysdatabases where name=\’" _
                        & (database & "\’"))
            Dim bRet As Boolean = False
            Using sqlConnection As SqlConnection = New SqlConnection(connString)
                sqlConnection.Open()
                Using sqlCmd As SqlCommand = New SqlCommand(cmdText, sqlConnection)
                    Using reader As SqlDataReader = sqlCmd.ExecuteReader
                        bRet = reader.HasRows
                    End Using
                End Using
            End Using
            CheckDatabaseExists = bRet
        Catch ex As Exception
            CheckDatabaseExists = False

        End Try
    End Function


    Public Function ExecuteSQLCommand(SQLstr As String, Table As String) As Integer
        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder(Table)
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If

        Catch ex As Exception
            ' MessageBox.Show("Error while Executing SQL Command ..." & ex.Message, "Execute SQL Command")

            Return 0
        End Try
        Return 1
    End Function

    Public Function ExecuteSQLQuery(SQLstr As String, Table As String) As String
        Dim ReadObject As Object
        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    ReadObject = dr.Item(0)
                    Return ReadObject
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder(Table)
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    ReadObject = drLocal.Item(0)
                    Return ReadObject
                End While
                atsLocal.Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Error while Executing SQL Query ..." & ex.Message, "Execute SQL Query")
            Return 0

        End Try
        Return 1
    End Function

    Public Function GetSpecification(Test As String, Optional freq As Double = 0.0) As Object
        Dim SQLStr As String
        Dim Count As Integer = 0
        Try
            If Job = "" Or Job = Nothing Then
                MsgBox("Please select a Job")
                GetSpecification = "N/A"
            End If
            GetSpecification = 0.0
            If Test = "StartFreqMHz" And SpecStartFreq <> Nothing Then
                GetSpecification = SpecStartFreq
                GoTo SkipDataBase
            ElseIf Test = "StopFreqMHz" And SpecStopFreq <> Nothing And SpecStopFreq <> 0.0 Then
                GetSpecification = SpecStopFreq
                GoTo SkipDataBase
                GoTo DataBaseSearch
            ElseIf Test = "Quantity" And Quantity <> Nothing Then
                GetSpecification = Quantity
                GoTo SkipDataBase
            ElseIf Test = "Title" And Title <> Nothing Then
                GetSpecification = Title
                GoTo SkipDataBase
            ElseIf Test = "InsertionLoss" And SpecIL <> Nothing Then
                GetSpecification = TruncateDecimal(SpecIL, 2)
                GoTo SkipDataBase
            ElseIf Test = "InsertionLoss1" And SpecIL <> Nothing Then
                GetSpecification = TruncateDecimal(SpecIL, 2)
                GoTo SkipDataBase
            ElseIf Test = "InsertionLoss2" And SpecIL_exp <> Nothing Then
                GetSpecification = TruncateDecimal(SpecIL_exp, 2)
                GoTo SkipDataBase
            ElseIf Test = "IL_ex" And SpecIL_exp <> Nothing Then
                GetSpecification = TruncateDecimal(SpecIL_exp, 2)
                GoTo SkipDataBase
            ElseIf Test = "VSWR" And SpecRL <> Nothing Then
                GetSpecification = TruncateDecimal(SpecRL, 1)
                GoTo SkipDataBase
            ElseIf Test = "Iso" Or Test = "Isolation" And SpecISO <> Nothing Then
                GetSpecification = TruncateDecimal(SpecISO, 1)
                GoTo SkipDataBase
            ElseIf Test = "IsolationL" And SpecISOL <> Nothing Then
                GetSpecification = TruncateDecimal(SpecISOL, 1)
                GoTo SkipDataBase
            ElseIf Test = "IsolationH" And SpecISOL <> Nothing Then
                GetSpecification = TruncateDecimal(SpecISOH, 1)
                GoTo SkipDataBase
            ElseIf Test = "AmplitudeBalance" And SpecAB <> Nothing Then
                GetSpecification = TruncateDecimal(SpecAB, 2)
                If Not SpecAB_TF Then GoTo SkipDataBase
            ElseIf Test = "PhaseBalance" And SpecPB <> Nothing Then
                GetSpecification = TruncateDecimal(SpecPB, 1)
                GoTo SkipDataBase
            ElseIf Test = "Coupling" And SpecCOUP <> Nothing Then
                GetSpecification = TruncateDecimal(SpecCOUP, 1)
                GoTo SkipDataBase
            ElseIf Test = "CoupPlusMinus" And SpecCOUPPM <> Nothing Then
                GetSpecification = TruncateDecimal(SpecCOUPPM, 1)
                GoTo SkipDataBase
            ElseIf Test = "Directivity" And SpecDIRECT <> Nothing Then
                GetSpecification = TruncateDecimal(SpecDIRECT, 1)
                GoTo SkipDataBase
            ElseIf Test = "CoupledFlatness" And SpecCOUPFLAT <> Nothing Then
                GetSpecification = TruncateDecimal(SpecCOUPFLAT, 2)
                GoTo SkipDataBase
            ElseIf Test = "Ports" And SpecPorts <> Nothing Then
                GetSpecification = SpecPorts
                GoTo SkipDataBase
            ElseIf Test = "PartsPerHour" And PPH <> Nothing Then
                GetSpecification = PPH
                GoTo SkipDataBase
            ElseIf Test = "SwitchPorts" And SwitchPorts <> Nothing Then
                GetSpecification = SwitchPorts
                GoTo SkipDataBase
                'ElseIf Test = "Offset1" And Offset1 <> Nothing Then
                '    GetSpecification = Offset1
                '    GoTo SkipDataBase
                'ElseIf Test = "Offset2" And Offset2 <> Nothing Then
                '    GetSpecification = Offset2
                '    GoTo SkipDataBase
                'ElseIf Test = "Offset3" And Offset3 <> Nothing Then
                '    GetSpecification = Offset3
                '    GoTo SkipDataBase
                'ElseIf Test = "Offset4" And Offset4 <> Nothing Then
                '    GetSpecification = Offset4
                '    GoTo SkipDataBase
                'ElseIf Test = "Offset5" And Offset5 <> Nothing Then
                '    GetSpecification = Offset5
                '    GoTo SkipDataBase
                'ElseIf Test = "Test1" And Test1 <> Nothing Then
                '    GetSpecification = Test1
                '    GoTo SkipDataBase
                'ElseIf Test = "Test2" And Test2 <> Nothing Then
                '    GetSpecification = Test2
                '    GoTo SkipDataBase
                'ElseIf Test = "Test3" And Test3 <> Nothing Then
                '    GetSpecification = Test3
                '    GoTo SkipDataBase
                'ElseIf Test = "Test4" And Test4 <> Nothing Then
                '    GetSpecification = Test4
                '    GoTo SkipDataBase
                'ElseIf Test = "Test5" And Test5 <> Nothing Then
                '    GetSpecification = Test5
                '    GoTo SkipDataBase
            ElseIf Test = "SpecType" And SpecType <> Nothing And SpecType <> "0" Then
                GetSpecification = SpecType
                GoTo SkipDataBase
            Else
                GetSpecification = 0.0
                GoTo DataBaseSearch
            End If
DataBaseSearch:

            Count = 0
            If Job = "" Then Job = frmAUTOTEST.cmbJob.Text
            GetSpecification = 0
Retry:
            SQLStr = "SELECT * from Specifications where JobNumber = '" & Job & "'"
            GetSpecification = 0.0
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If Not IsDBNull(dr.Item(0)) Then SpecID = CStr(dr.Item(0))
                    If Not IsDBNull(dr.Item(1)) Then SpecType = CStr(dr.Item(1))
                    If Not IsDBNull(dr.Item(4)) Then Title = CStr(dr.Item(4))
                    If Not IsDBNull(dr.Item(5)) Then Quantity = CInt(dr.Item(5))
                    If Not IsDBNull(dr.Item(6)) Then SpecStartFreq = CDbl(dr.Item(6))
                    If Not IsDBNull(dr.Item(7)) Then SpecStopFreq = CDbl(dr.Item(7))
                    If Not IsDBNull(dr.Item(8)) Then SpecCuttoffFreq = CDbl(dr.Item(8))
                    If Not IsDBNull(dr.Item(11)) Then SpecIL = TruncateDecimal(CDbl(dr.Item(11)), 2)
                    If Not IsDBNull(dr.Item(10)) Then SpecRL = TruncateDecimal(VSWRtoRL(CDbl(dr.Item(10))), 1)
                    If Not IsDBNull(dr.Item(12)) Then SpecISO = TruncateDecimal(CDbl(dr.Item(12)), 1)
                    If Not IsDBNull(dr.Item(12)) Then SpecISO = TruncateDecimal(CDbl(dr.Item(12)), 1)
                    If Not IsDBNull(dr.Item(12)) Then SpecISOL = TruncateDecimal(CDbl(dr.Item(12)), 1)
                    If Not IsDBNull(dr.Item(13)) Then SpecISOH = TruncateDecimal(CDbl(dr.Item(13)), 1)
                    If Not IsDBNull(dr.Item(14)) Then SpecAB = TruncateDecimal(CDbl(dr.Item(14)), 2)
                    If Not IsDBNull(dr.Item(18)) Then SpecPB = TruncateDecimal(CDbl(dr.Item(18)), 1)
                    If Not IsDBNull(dr.Item(15)) Then SpecCOUP = TruncateDecimal(CDbl(dr.Item(15)), 1)
                    If Not IsDBNull(dr.Item(16)) Then SpecCOUPPM = TruncateDecimal(CDbl(dr.Item(16)), 1)
                    If Not IsDBNull(dr.Item(17)) Then SpecDIRECT = TruncateDecimal(CDbl(dr.Item(17)), 1)
                    If Not IsDBNull(dr.Item(19)) Then SpecCOUPFLAT = TruncateDecimal(CDbl(dr.Item(19)), 2)
                    If Not IsDBNull(dr.Item(9)) Then SpecPorts = CInt(dr.Item(9))
                    If Not IsDBNull(dr.Item(32)) Then PPH = CInt(dr.Item(32))
                    If Not IsDBNull(dr.Item(22)) Then Offset1 = CInt(dr.Item(22))
                    If Not IsDBNull(dr.Item(23)) Then Offset2 = CInt(dr.Item(23))
                    If Not IsDBNull(dr.Item(24)) Then Offset3 = CInt(dr.Item(24))
                    If Not IsDBNull(dr.Item(25)) Then Offset4 = CInt(dr.Item(25))
                    If Not IsDBNull(dr.Item(26)) Then Offset5 = CInt(dr.Item(26))
                    If Not IsDBNull(dr.Item(27)) Then Test1 = CInt(dr.Item(27))
                    If Not IsDBNull(dr.Item(28)) Then Test2 = CInt(dr.Item(28))
                    If Not IsDBNull(dr.Item(29)) Then Test3 = CInt(dr.Item(29))
                    If Not IsDBNull(dr.Item(30)) Then Test4 = CInt(dr.Item(30))
                    If Not IsDBNull(dr.Item(31)) Then Test5 = CInt(dr.Item(31))
                    If Not IsDBNull(dr.GetValue(47)) Then
                        If dr.Item(47) = 1 Then
                            SpecAB_TF = True
                            If Not IsDBNull(dr.GetValue(42)) Then SpecAB_exp = dr.Item(42)
                            If Not IsDBNull(dr.GetValue(43)) Then SpecAB_start1 = dr.Item(43)
                            If Not IsDBNull(dr.GetValue(44)) Then SpecAB_start2 = dr.Item(44)
                            If Not IsDBNull(dr.GetValue(45)) Then SpecAB_stop1 = dr.Item(45)
                            If Not IsDBNull(dr.GetValue(46)) Then SpecAB_stop2 = dr.Item(46)
                        Else
                            SpecAB_TF = False
                        End If
                    End If
                    If Not IsDBNull(dr.GetValue(91)) Then
                        If dr.Item(91) Then
                            IL_TF = True
                            If Not IsDBNull(dr.GetValue(92)) Then SpecIL_exp = dr.Item(92)
                            If Not IsDBNull(dr.GetValue(93)) Then SpecIL_start1 = dr.Item(93)
                            If Not IsDBNull(dr.GetValue(94)) Then SpecIL_start2 = dr.Item(94)
                            If Not IsDBNull(dr.GetValue(95)) Then SpecIL_stop1 = dr.Item(95)
                            If Not IsDBNull(dr.GetValue(96)) Then SpecIL_stop2 = dr.Item(96)
                        Else
                            IL_TF = False
                        End If
                    End If
                    If Not IsDBNull(dr.GetValue(100)) Then SwitchPorts = dr.Item(100)
                    If Test = "SpecType" Then If Not IsDBNull(dr.Item(1)) Then GetSpecification = CDbl(dr.Item(1))
                    If Test = "Title" Then If Not IsDBNull(dr.Item(4)) Then GetSpecification = CDbl(dr.Item(4))
                    If Test = "Quantity" Then If Not IsDBNull(dr.Item(5)) Then GetSpecification = CDbl(dr.Item(5))
                    If Test = "StartFreqMHz" Then If Not IsDBNull(dr.Item(6)) Then GetSpecification = CDbl(dr.Item(6))
                    If Test = "StopFreqMHz" Then If Not IsDBNull(dr.Item(7)) Then GetSpecification = CDbl(dr.Item(7))
                    If Test = "InsertionLoss" Then If Not IsDBNull(dr.Item(11)) Then GetSpecification = TruncateDecimal(CDbl(dr.Item(11)), 2)
                    If Test = "InsertionLoss1" Then If Not IsDBNull(dr.Item(11)) Then GetSpecification = TruncateDecimal(CDbl(dr.Item(11)), 2)
                    If Test = "IL_ex" Then If Not IsDBNull(dr.Item(91)) Then GetSpecification = TruncateDecimal(CDbl(dr.Item(91)), 2)
                    If Test = "VSWR" Then If Not IsDBNull(dr.Item(10)) Then GetSpecification = TruncateDecimal(Math.Round(VSWRtoRL(CDbl(dr.Item(10)))), 1)
                    If Test = "Isolation" Or Test = "Iso" Then If Not IsDBNull(dr.Item(12)) Then GetSpecification = Math.Round(CDbl(dr.Item(12)), 1)
                    If Test = "Isolation2" Then If Not IsDBNull(dr.Item(13)) Then GetSpecification = TruncateDecimal(CDbl(dr.Item(13)), 1)
                    If Test = "CutOffFreqMHz" Then If Not IsDBNull(dr.Item(8)) Then GetSpecification = CDbl(dr.Item(8))
                    If Test = "IsolationL" And Not IsDBNull(dr.Item(12)) Then GetSpecification = CDbl(dr.Item(12))
                    If Test = "IsolationH" And Not IsDBNull(dr.Item(13)) Then GetSpecification = CDbl(dr.Item(13))
                    If Test = "AmplitudeBalance" Then If Not IsDBNull(dr.Item(14)) Then GetSpecification = TruncateDecimal(CDbl(dr.Item(14)), 2)
                    If Test = "PhaseBalance" Then If Not IsDBNull(dr.Item(18)) Then GetSpecification = TruncateDecimal(CDbl(dr.Item(18)), 1)
                    If Test = "Coupling" Then If Not IsDBNull(dr.Item(15)) Then GetSpecification = TruncateDecimal(CDbl(dr.Item(15)), 1)
                    If Test = "CouplingPM" Then If Not IsDBNull(dr.Item(16)) Then GetSpecification = TruncateDecimal(CDbl(dr.Item(16)), 1)
                    If Test = "Directivity" Then If Not IsDBNull(dr.Item(17)) Then GetSpecification = TruncateDecimal(CDbl(dr.Item(17)), 1)
                    If Test = "CoupledFlatness" Then If Not IsDBNull(dr.Item(19)) Then GetSpecification = TruncateDecimal(CDbl(dr.Item(19)), 2)
                    If Test = "SpecID" Then If Not IsDBNull(dr.Item(0)) Then GetSpecification = CDbl(dr.Item(0))
                    If Test = "Ports" Then If Not IsDBNull(dr.Item(9)) Then GetSpecification = CDbl(dr.Item(9))
                    If Test = "Offset1" Then If Not IsDBNull(dr.Item(22)) Then GetSpecification = CDbl(dr.Item(22))
                    If Test = "Offset2" Then If Not IsDBNull(dr.Item(23)) Then GetSpecification = CDbl(dr.Item(23))
                    If Test = "Offset3" Then If Not IsDBNull(dr.Item(24)) Then GetSpecification = CDbl(dr.Item(24))
                    If Test = "Offset4" Then If Not IsDBNull(dr.Item(25)) Then GetSpecification = CDbl(dr.Item(25))
                    If Test = "Offset5" Then If Not IsDBNull(dr.Item(26)) Then GetSpecification = CDbl(dr.Item(26))
                    If Test = "Test1" Then If Not IsDBNull(dr.Item(27)) Then GetSpecification = CDbl(dr.Item(27))
                    If Test = "Test2" Then If Not IsDBNull(dr.Item(28)) Then GetSpecification = CDbl(dr.Item(28))
                    If Test = "Test3" Then If Not IsDBNull(dr.Item(29)) Then GetSpecification = CDbl(dr.Item(29))
                    If Test = "Test4" Then If Not IsDBNull(dr.Item(30)) Then GetSpecification = CDbl(dr.Item(30))
                    If Test = "Test5" Then If Not IsDBNull(dr.Item(31)) Then GetSpecification = CDbl(dr.Item(31))
                    If Test = "PartsPerHour" Then If Not IsDBNull(dr.Item(32)) Then GetSpecification = CInt(dr.Item(32))
                    If Test = "SwitchPorts" Then If Not IsDBNull(dr.GetValue(100)) Then SwitchPorts = dr.Item(100)
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If Not IsDBNull(drLocal.Item(0)) Then SpecID = CStr(drLocal.Item(0))
                    If Not IsDBNull(drLocal.Item(1)) Then SpecType = CStr(drLocal.Item(1))
                    If Not IsDBNull(drLocal.Item(4)) Then Title = CStr(drLocal.Item(4))
                    If Not IsDBNull(drLocal.Item(5)) Then Quantity = CInt(drLocal.Item(5))
                    If Not IsDBNull(drLocal.Item(6)) Then SpecStartFreq = CDbl(drLocal.Item(6))
                    If Not IsDBNull(drLocal.Item(7)) Then SpecStopFreq = CDbl(drLocal.Item(7))
                    If Not IsDBNull(drLocal.Item(8)) Then SpecCuttoffFreq = CDbl(drLocal.Item(8))
                    If Not IsDBNull(drLocal.Item(11)) Then SpecIL = TruncateDecimal(CDbl(drLocal.Item(11)), 2)
                    If Not IsDBNull(drLocal.Item(10)) Then SpecRL = TruncateDecimal(VSWRtoRL(CDbl(drLocal.Item(10))), 1)
                    If Not IsDBNull(drLocal.Item(12)) Then SpecISO = TruncateDecimal(CDbl(drLocal.Item(12)), 1)
                    If Not IsDBNull(drLocal.Item(12)) Then SpecISOL = TruncateDecimal(CDbl(drLocal.Item(12)), 1)
                    If Not IsDBNull(drLocal.Item(13)) Then SpecISOH = TruncateDecimal(CDbl(drLocal.Item(13)), 1)
                    If Not IsDBNull(drLocal.Item(14)) Then SpecAB = TruncateDecimal(CDbl(drLocal.Item(14)), 2)
                    If Not IsDBNull(drLocal.Item(18)) Then SpecPB = TruncateDecimal(CDbl(drLocal.Item(18)), 1)
                    If Not IsDBNull(drLocal.Item(15)) Then SpecCOUP = TruncateDecimal(CDbl(drLocal.Item(15)), 1)
                    If Not IsDBNull(drLocal.Item(16)) Then SpecCOUPPM = TruncateDecimal(CDbl(drLocal.Item(16)), 1)
                    If Not IsDBNull(drLocal.Item(17)) Then SpecDIRECT = TruncateDecimal(CDbl(drLocal.Item(17)), 1)
                    If Not IsDBNull(drLocal.Item(19)) Then SpecCOUPFLAT = TruncateDecimal(CDbl(drLocal.Item(19)), 2)
                    If Not IsDBNull(drLocal.Item(9)) Then SpecPorts = CInt(drLocal.Item(9))
                    If Not IsDBNull(drLocal.Item(32)) Then PPH = CInt(drLocal.Item(32))


                    If Test = "SpecType" Then If Not IsDBNull(drLocal.Item(1)) Then GetSpecification = CDbl(drLocal.Item(1))
                    If Test = "Title" Then If Not IsDBNull(drLocal.Item(4)) Then GetSpecification = CDbl(drLocal.Item(4))
                    If Test = "Quantity" Then If Not IsDBNull(drLocal.Item(5)) Then GetSpecification = CDbl(drLocal.Item(5))
                    If Test = "StartFreqMHz" Then If Not IsDBNull(drLocal.Item(6)) Then GetSpecification = CDbl(drLocal.Item(6))
                    If Test = "StopFreqMHz" Then If Not IsDBNull(drLocal.Item(7)) Then GetSpecification = CDbl(drLocal.Item(7))
                    If Test = "InsertionLoss" Then If Not IsDBNull(drLocal.Item(12)) Then GetSpecification = TruncateDecimal(CDbl(drLocal.Item(11)), 2)
                    If Test = "VSWR" Then If Not IsDBNull(drLocal.Item(11)) Then GetSpecification = TruncateDecimal(Math.Round(VSWRtoRL(CDbl(drLocal.Item(10)))), 1)
                    If Test = "Isolation" Or Test = "Iso" Then If Not IsDBNull(drLocal.Item(12)) Then GetSpecification = Math.Round(CDbl(drLocal.Item(12)), 1)
                    If Test = "IsolationL" Then If Not IsDBNull(drLocal.Item(12)) Then GetSpecification = TruncateDecimal(CDbl(drLocal.Item(12)), 1)
                    If Test = "IsolationH" Then If Not IsDBNull(drLocal.Item(13)) Then GetSpecification = TruncateDecimal(CDbl(drLocal.Item(13)), 1)
                    If Test = "CutOffFreqMHz" Then If Not IsDBNull(drLocal.Item(8)) Then GetSpecification = TruncateDecimal(CDbl(drLocal.Item(8)), 1)
                    If Test = "AmplitudeBalance" Then If Not IsDBNull(drLocal.Item(14)) Then GetSpecification = TruncateDecimal(CDbl(drLocal.Item(14)), 2)
                    If Test = "PhaseBalance" Then If Not IsDBNull(drLocal.Item(18)) Then GetSpecification = TruncateDecimal(CDbl(drLocal.Item(18)), 1)
                    If Test = "Coupling" Then If Not IsDBNull(drLocal.Item(15)) Then GetSpecification = TruncateDecimal(CDbl(drLocal.Item(15)), 1)
                    If Test = "CouplingPM" Then If Not IsDBNull(drLocal.Item(16)) Then GetSpecification = TruncateDecimal(CDbl(drLocal.Item(16)), 1)
                    If Test = "Directivity" Then If Not IsDBNull(drLocal.Item(17)) Then GetSpecification = TruncateDecimal(CDbl(drLocal.Item(17)), 1)
                    If Test = "CoupledFlatness" Then If Not IsDBNull(drLocal.Item(19)) Then GetSpecification = TruncateDecimal(CDbl(drLocal.Item(19)), 2)
                    If Test = "SpecID" Then If Not IsDBNull(drLocal.Item(0)) Then GetSpecification = CDbl(drLocal.Item(0))
                    If Test = "Ports" Then If Not IsDBNull(drLocal.Item(9)) Then GetSpecification = CDbl(drLocal.Item(9))
                    If Test = "Offset1" Then If Not IsDBNull(drLocal.Item(22)) Then GetSpecification = CDbl(drLocal.Item(22))
                    If Test = "Offset2" Then If Not IsDBNull(drLocal.Item(23)) Then GetSpecification = CDbl(drLocal.Item(23))
                    If Test = "Offset3" Then If Not IsDBNull(drLocal.Item(24)) Then GetSpecification = CDbl(drLocal.Item(24))
                    If Test = "Offset4" Then If Not IsDBNull(drLocal.Item(25)) Then GetSpecification = CDbl(drLocal.Item(25))
                    If Test = "Offset5" Then If Not IsDBNull(drLocal.Item(25)) Then GetSpecification = CDbl(drLocal.Item(26))
                    If Test = "Test1" Then If Not IsDBNull(drLocal.Item(27)) Then GetSpecification = CDbl(drLocal.Item(27))
                    If Test = "Test2" Then If Not IsDBNull(drLocal.Item(28)) Then GetSpecification = CDbl(drLocal.Item(28))
                    If Test = "Test3" Then If Not IsDBNull(drLocal.Item(29)) Then GetSpecification = CDbl(drLocal.Item(29))
                    If Test = "Test4" Then If Not IsDBNull(drLocal.Item(30)) Then GetSpecification = CDbl(drLocal.Item(30))
                    If Test = "Test5" Then If Not IsDBNull(drLocal.Item(31)) Then GetSpecification = CDbl(drLocal.Item(31))
                    If Test = "PartsPerHour" Then If Not IsDBNull(drLocal.Item(32)) Then GetSpecification = CInt(drLocal.Item(32))
                End While
                atsLocal.Close()
            End If
            If SpecCuttoffFreq <> 0 Then
                ISO_TF = True
            End If
            '*******************************************************************************
        Catch
            GetSpecification = 0
        End Try
SkipDataBase:

    End Function

    Public Function CheckforRow(SQLStr As String, Table As String) As Integer
        Try
            Dim CountRow As Integer = 0

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    CountRow = CountRow + 1
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder(Table)
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    CountRow = CountRow + 1
                End While
                atsLocal.Close()
            End If
            Return CountRow
        Catch
            Return 0
        End Try
    End Function

    Public Function GetLoss() As Double

        Dim SQLstr As String
        Dim Temp() As String
        Try
            GetLoss = 0
            SQLstr = "SELECT * from PortConfig where PartNumber = '" & Part & "'"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    Temp = Split(dr.Item(5), "dB")

                    If Temp(0).Contains("J") Then
                        GetLoss = 0
                    Else
                        GetLoss = CDbl(Temp(0))
                    End If
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    Temp = Split(drLocal.Item(5), "dB")

                    If InStr(Temp(0), "J") Then
                        GetLoss = 0
                    Else
                        GetLoss = CDbl(Temp(0))
                    End If
                End While
                atsLocal.Close()
            End If

        Catch ex As Exception
            GetLoss = 0

        End Try

    End Function

    Public Function GetVNAFreq() As Double
        Try
            Dim CountRow As Integer = 0
            Dim SQLstr As String

            SQLstr = "select * from WorkStation where ComputerName = '" & GetComputerName() & "'"
            GetVNAFreq = 0
            'Job = frmAUTOTEST.cmbJob.Text
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(5) IsNot Nothing Then GetVNAFreq = CDbl(dr.Item(5))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(5) IsNot Nothing Then GetVNAFreq = CDbl(drLocal.Item(5))
                End While
                atsLocal.Close()
            End If
        Catch
            Return 0
        End Try
    End Function

    Public Sub SetVNAFreq(freq As String)
        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE WorkStation Set VNAFreq = '" & freq & "' where ComputerName = '" & GetComputerName() & "'"
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.CommandText = "UPDATE WorkStation Set VNAFreq = '" & freq & "' where ComputerName = '" & GetComputerName() & "'"""
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Sub SaveTestData(Test As String, Value As Double)
        Dim SQLStr As String

        If TweakMode Then Exit Sub
        Try
            SQLStr = "SELECT * from TestData where JobNumber = '" & Job & "' And SerialNumber = '" & SerialNumber & "' and WorkStation = '" & GetComputerName() & "' and artwork_rev = '" & ArtworkRevision & "'"
            If SQL.CheckforRow(SQLStr, "NetworkData") = 0 Then
                SQLStr = "Insert Into TestData (JobNumber, PartNumber,SerialNumber,WorkStation,artwork_rev) values ('" & Job & "','" & Part & "','" & SerialNumber & "','" & GetComputerName() & "','" & ArtworkRevision & "')"
                SQL.ExecuteSQLCommand(SQLStr, "NetworkData")
            End If

            SQLStr = "UPDATE TestData Set " & Test & " = '" & Value & "' where JobNumber = '" & Job & "' And SerialNumber = '" & SerialNumber & "' and WorkStation = '" & GetComputerName() & "' and artwork_rev = '" & ArtworkRevision & "'"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkData")

            SQLStr = "UPDATE TestData Set artwork_rev  = '" & ArtworkRevision & "' where JobNumber = '" & Job & "' And SerialNumber = '" & SerialNumber & "' and WorkStation = '" & GetComputerName() & "' and artwork_rev = '" & ArtworkRevision & "'"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkData")

            SQLStr = "UPDATE TestData Set Operator  = '" & User & "' where JobNumber = '" & Job & "' And SerialNumber = '" & SerialNumber & "' and WorkStation = '" & GetComputerName() & "' and artwork_rev = '" & ArtworkRevision & "'"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkData")

        Catch ex As Exception

        End Try
    End Sub
    Public Sub CleanUpEffeciency()

        Dim SQLstr As String
        Dim TempNow(4) As String
        Dim TempThen(4) As String
        Try
            If TweakMode Then GoTo PAUSED2

            TempNow = Split(Now.ToShortDateString, "/")
            SQLstr = "SELECT * from Effeciency Where WorkStation = '" & GetComputerName() & "'"

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    TempThen = Split(dr.Item(5), "/")
                    If CInt(TempThen(0)) <= CInt(TempNow(0)) And CInt(TempThen(1)) <= CInt(TempNow(1)) And CInt(TempThen(2)) < CInt(TempNow(2)) Then GoTo Archived
                    If IsDBNull(dr.Item(1)) Or IsDBNull(dr.Item(2)) Or IsDBNull(dr.Item(3)) Or IsDBNull(dr.Item(4)) Or IsDBNull(dr.Item(5)) Or IsDBNull(dr.Item(6)) _
                         Or IsDBNull(dr.Item(7)) Or IsDBNull(dr.Item(8)) Or IsDBNull(dr.Item(9)) Then GoTo Delete
                    If CStr(dr.Item(9)).Contains("Complete") Then GoTo IGNORE
                    If CStr(dr.Item(9)).Contains("Ready") Then GoTo IGNORE
                    If dr.Item(5) = Now.ToShortDateString Then GoTo IGNORE
                    If CInt(TempThen(2)) < CInt(TempNow(2)) Then GoTo COMPLETE
                    If CInt(TempThen(2)) = CInt(TempNow(2)) And CInt(TempThen(0)) < CInt(TempNow(0)) Then GoTo COMPLETE
                    If CInt(TempThen(2)) = CInt(TempThen(2)) And CInt(TempThen(0)) = CInt(TempNow(0)) And Math.Abs(CInt(TempThen(1)) - CInt(TempNow(1))) > 10 Then GoTo COMPLETE
                    If CInt(TempThen(2)) = CInt(TempThen(2)) And CInt(TempThen(0)) = CInt(TempNow(0)) And Math.Abs(CInt(TempThen(1)) - CInt(TempNow(1))) > 2 Then GoTo COMPLETE
                    If CInt(TempThen(2)) = CInt(TempThen(2)) And CInt(TempThen(0)) = CInt(TempNow(0)) And Math.Abs(CInt(TempThen(1)) - CInt(TempNow(1))) < 2 Then GoTo PAUSED
PAUSED:
                    SQLstr = "UPDATE Effeciency set RunStatus = 'Paused' Where ActiveDate = '" & dr.Item(5) & "' and WorkStation = '" & GetComputerName() & "'"
                    SQL.ExecuteSQLCommand(SQLstr, "Effeciency")
                    GoTo Ignore
COMPLETE:
                    SQLstr = "UPDATE Effeciency set RunStatus = 'Complete' Where ActiveDate = '" & dr.Item(5) & "' and WorkStation = '" & GetComputerName() & "'"
                    SQL.ExecuteSQLCommand(SQLstr, "Effeciency")
                    GoTo Ignore
Archived:
                    SQLstr = "UPDATE Effeciency set RunStatus = 'Archived' Where ActiveDate = '" & dr.Item(5) & "' and WorkStation = '" & GetComputerName() & "'"
                    SQL.ExecuteSQLCommand(SQLstr, "Effeciency")

Delete:
                    SQLstr = "Delete Effeciency Where ID = " & dr.Item(0)
                    SQL.ExecuteSQLCommand(SQLstr, "Effeciency")
IGNORE:
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("Effeciency")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    TempThen = Split(drLocal.Item(5), "/")
                    If CInt(TempThen(0)) <= CInt(TempNow(0)) And CInt(TempThen(1)) <= CInt(TempNow(1)) And CInt(TempThen(2)) < CInt(TempNow(2)) Then GoTo Archived2
                    If CStr(drLocal.Item(9)).Contains("Complete") Then GoTo IGNORE2
                    If CStr(drLocal.Item(9)).Contains("Ready") Then GoTo IGNORE2
                    If drLocal.Item(5) = Now.ToShortDateString Then GoTo IGNORE2
                    If CInt(TempThen(2)) < CInt(TempNow(2)) Then GoTo COMPLETE2
                    If CInt(TempThen(2)) = CInt(TempNow(2)) And CInt(TempThen(0)) < CInt(TempNow(0)) Then GoTo COMPLETE2
                    If CInt(TempThen(2)) = CInt(TempThen(2)) And CInt(TempThen(0)) = CInt(TempNow(0)) And Math.Abs(CInt(TempThen(1)) - CInt(TempNow(1))) > 10 Then GoTo COMPLETE2
                    If CInt(TempThen(2)) = CInt(TempThen(2)) And CInt(TempThen(0)) = CInt(TempNow(0)) And Math.Abs(CInt(TempThen(1)) - CInt(TempNow(1))) > 2 Then GoTo COMPLETE2
                    If CInt(TempThen(2)) = CInt(TempThen(2)) And CInt(TempThen(0)) = CInt(TempNow(0)) And Math.Abs(CInt(TempThen(1)) - CInt(TempNow(1))) < 2 Then GoTo PAUSED2

PAUSED2:
                    SQLstr = "UPDATE Effeciency set RunStatus = 'Paused' Where ActiveDate = '" & drLocal.Item(5) & "' and WorkStation = '" & GetComputerName() & "'"
                    SQL.ExecuteSQLCommand(SQLstr, "Effeciency")
                    GoTo IGNORE2
COMPLETE2:
                    SQLstr = "UPDATE Effeciency set RunStatus = 'Complete' Where ActiveDate = '" & drLocal.Item(5) & "' and WorkStation = '" & GetComputerName() & "'"
                    SQL.ExecuteSQLCommand(SQLstr, "Effeciency")
                    GoTo Ignore2
Archived2:
                    SQLstr = "UPDATE Effeciency set RunStatus = 'Archived' Where ActiveDate = '" & drLocal.Item(5) & "' and WorkStation = '" & GetComputerName() & "'"
                    SQL.ExecuteSQLCommand(SQLstr, "Effeciency")
IGNORE2:
                End While
                drLocal.Close()
            End If

        Catch ex As Exception
        End Try

    End Sub
    Public Sub UpdateEffeciency(runStatus As String, Effeciency As String, ActiveDate As String, UUTNum As Integer)
        Dim SQLStr As String
        Try
            If Effeciency = "N/A" Then Exit Sub
            SQLStr = "SELECT * from Effeciency where JobNumber = '" & Job & "' And WorkStation = '" & GetComputerName() & "' And RunStatus <> 'Complete' And RunStatus <> 'Archived'"
            If SQL.CheckforRow(SQLStr, "Effeciency") = 0 Then
                SQLStr = "Insert Into Effeciency (JobNumber, PartNumber, RunStatus, WorkStation, TotalUUTs) values ('" & Job & "','" & Part & "','" & runStatus & "','" & GetComputerName() & "'," & Quantity & ")"
                SQL.ExecuteSQLCommand(SQLStr, "Effeciency")
            End If

            SQLStr = "UPDATE Effeciency Set CompleteUUTs = " & UUTNum & " where JobNumber = '" & Job & "' And WorkStation = '" & GetComputerName() & "' And RunStatus <> 'Complete' And RunStatus <> 'Archived'"
            SQL.ExecuteSQLCommand(SQLStr, "Effeciency")

            SQLStr = "UPDATE Effeciency Set ActiveDate = '" & Now.Date.ToShortDateString & "'  where JobNumber = '" & Job & "' And WorkStation = '" & GetComputerName() & "' And RunStatus <> 'Complete' And RunStatus <> 'Archived'"
            SQL.ExecuteSQLCommand(SQLStr, "Effeciency")

            SQLStr = "UPDATE Effeciency Set EffeciencyStatus = '" & Effeciency & "' where JobNumber = '" & Job & "' And WorkStation = '" & GetComputerName() & "' And RunStatus <> 'Complete' And RunStatus <> 'Archived'"
            SQL.ExecuteSQLCommand(SQLStr, "Effeciency")

            SQLStr = "UPDATE Effeciency Set RunStatus = '" & runStatus & "' where JobNumber = '" & Job & "' And WorkStation = '" & GetComputerName() & "' And RunStatus <> 'Complete' And RunStatus <> 'Archived'"
            SQL.ExecuteSQLCommand(SQLStr, "Effeciency")

            SQLStr = "UPDATE Effeciency Set Operator = '" & User & "' where JobNumber = '" & Job & "' And WorkStation = '" & GetComputerName() & "' And RunStatus <> 'Complete' And RunStatus <> 'Archived'"
            SQL.ExecuteSQLCommand(SQLStr, "Effeciency")

        Catch ex As Exception

        End Try

    End Sub
    Public Sub SaveFailureLog(Value As String)
        Dim SQLStr As String
        Try
            SQLStr = "SELECT * from TestData  where JobNumber = '" & Job & "' And SerialNumber = '" & SerialNumber & "'"
            If SQL.CheckforRow(SQLStr, "NetworkData") = 0 Then
                SQLStr = "Insert Into TestData (JobNumber, PartNumber,SerialNumber,WorkStation) values ('" & Job & "','" & Part & "','" & SerialNumber & "','" & GetComputerName() & "')"
                SQL.ExecuteSQLCommand(SQLStr, "NetworkData")
            End If

            SQLStr = "UPDATE TestData Set FailureLog = '" & Value & "' where JobNumber = '" & Job & "' And SerialNumber = '" & SerialNumber & "' and WorkStation = '" & GetComputerName() & "'"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkData")


        Catch ex As Exception

        End Try

    End Sub
    Public Sub SaveUser(Value As String)
        Dim SQLStr As String
        Try
            SQLStr = "SELECT * from Workstation where ComputerName = '" & GetComputerName() & "'"
            If SQL.CheckforRow(SQLStr, "NetworkSpecs") = 0 Then
                SQLStr = "Insert Into Workstation (ComputerName,WorkStationName) values ('" & GetComputerName() & "','" & GetComputerName() & "')"
                SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
            End If

            SQLStr = "UPDATE Workstation Set Operator = '" & Value & "' where WorkStationName = '" & GetComputerName() & "'"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")


        Catch ex As Exception

        End Try

    End Sub
    Public Function GetBypass() As Integer
        Try
            Dim CountRow As Integer = 0
            GetBypass = 0
            Dim SQLstr As String = "SELECT * from Specifications where PartNumber = '" & PartSpec & "' And JobNumber = '" & jobSpec & "'"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(35) IsNot Nothing Then GetBypass = CStr(dr.Item(35))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(35) IsNot Nothing Then GetBypass = CStr(drLocal.Item(35))
                End While

                atsLocal.Close()
            End If
        Catch
            GetBypass = 0
        End Try
    End Function
    Public Function GetFailPercent() As Double
        Try
            Dim CountRow As Integer = 0
            GetFailPercent = 15
            Dim SQLstr As String = "SELECT * from Specifications where PartNumber = '" & PartSpec & "' And JobNumber = '" & jobSpec & "'"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(40) IsNot Nothing Then GetFailPercent = CDbl(dr.Item(40))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(40) IsNot Nothing Then GetFailPercent = CDbl(drLocal.Item(40))
                End While

                atsLocal.Close()
            End If
        Catch
            GetFailPercent = 15
        End Try
    End Function
    Public Sub SaveBypass(Value As Integer)
        Dim SQLStr As String
        Try
            SQLStr = "UPDATE Specifications Set Bypass = '" & Value & "' where PartNumber = '" & PartSpec & "' And JobNumber = '" & jobSpec & "'"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
        Catch ex As Exception

        End Try
    End Sub
    Public Sub SavePassword(Value As String)
        Dim SQLStr As String
        Try
            SQLStr = "UPDATE Workstation Set Password = '" & Value & "' where WorkStationName = 'ALLSTATIONS'"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")
        Catch ex As Exception

        End Try
    End Sub
    Public Function GetPassword() As String
        Try
            Dim CountRow As Integer = 0
            GetPassword = ""
            Dim SQLstr As String = "SELECT * from Workstation where WorkStationName = 'ALLSTATIONS'"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(5) IsNot Nothing Then GetPassword = CStr(dr.Item(6))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(0) IsNot Nothing Then GetPassword = CStr(drLocal.Item(6))
                End While

                atsLocal.Close()
            End If
        Catch
            GetPassword = ""
        End Try
    End Function
    Public Function GetreportStatus() As String
        Try
            Dim CountRow As Integer = 0
            Dim SQLstr As String = "SELECT * from ReportQueue where WorkStation = '" & GetComputerName() & "' and ReportStatus = 'test running'"
            Dim reportJobs(5) As String
            GetreportStatus = "no status"
            Dim x As Integer = 0
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(3) IsNot Nothing Then GetreportStatus = CStr(dr.Item(3))
                    If dr.Item(4) IsNot Nothing Then ReportJob = CStr(dr.Item(4))
                    reportJobs(x) = ReportJob
                    x += 1
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(3) IsNot Nothing Then GetreportStatus = CStr(drLocal.Item(3))
                    If drLocal.Item(4) IsNot Nothing Then ReportJob = CStr(drLocal.Item(4))
                    reportJobs(x) = ReportJob
                    x += 1
                End While

                atsLocal.Close()
            End If
        Catch
            GetreportStatus = ""
        End Try
    End Function
    Public Function GetLastUser() As String
        Try
            Dim CountRow As Integer = 0
            Dim SQLstr As String = "SELECT * from Workstation where WorkStationName = '" & GetComputerName() & "'"
            GetLastUser = ""
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(4) IsNot Nothing Then GetLastUser = CStr(dr.Item(4))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(4) IsNot Nothing Then GetLastUser = CStr(drLocal.Item(4))
                End While

                atsLocal.Close()
            End If
        Catch
            GetLastUser = ""
        End Try
    End Function

    Public Function GetTestID(SQLstr As String, Table As String) As Integer
        Try
            Dim CountRow As Integer = 0
            GetTestID = 0
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(5) IsNot Nothing Then GetTestID = CInt(dr.Item(0))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder(Table)
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(0) IsNot Nothing Then GetTestID = CInt(drLocal.Item(0))
                End While

                atsLocal.Close()
            End If

        Catch
            Return 0
        End Try
    End Function
    
    Public Function GetTraceID(SQLstr As String, Table As String) As Integer
        Try
            Dim CountRow As Integer = 0
            GetTraceID = 0
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(5) IsNot Nothing Then GetTraceID = CInt(dr.Item(0))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder(Table)
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(0) IsNot Nothing Then GetTraceID = CInt(drLocal.Item(0))
                End While

                atsLocal.Close()
            End If

        Catch
            Return 0
        End Try

    End Function
    Public Function GetTraceIDByTitle(Title As String, Optional Serial As String = "", Optional Job As String = "", Optional Station As String = "") As Integer
        Try
            GetTraceIDByTitle = 0
            If Serial = "" Then Serial = "UUT" & UUTNum_Reset
            If Job = "" Then Job = frmAUTOTEST.cmbJob.Text
            If Station = "" Then Station = WorkStation

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                'Dim cmd As SqlCommand = New SqlCommand("SELECT * from Trace where JobNumber =  '" & Job & "' and SerialNumber = '" & Serial & "' and Title = '" & Title & "' And Workstation = '" & GetComputerName() & "'", ats)
                Dim cmd As SqlCommand = New SqlCommand("SELECT * from Trace where JobNumber =  '" & Job & "' and SerialNumber = '" & Serial & "' and Title = '" & Title & "'", ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    GetTraceIDByTitle = CInt(dr.Item(0))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                'Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * from Trace where JobNumber =  '" & Job & "' and SerialNumber = '" & Serial & "' and Title = '" & Title & "' And Workstation = '" & GetComputerName() & "'", atsLocal)
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * from Trace where JobNumber =  '" & Job & "' and SerialNumber = '" & Serial & "' and Title = '" & Title & "'", atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    GetTraceIDByTitle = drLocal.Item(0)
                End While
                atsLocal.Close()
            End If
        Catch
            GetTraceIDByTitle = 0
        End Try

    End Function

    Public Function GetTitle() As String
        Try
            Dim CountRow As Integer = 0
            Dim SQLstr As String
            SQLstr = "Select * from Specifications where JobNumber ='" & Job & "'"

            GetTitle = ""
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(4) IsNot Nothing Then GetTitle = CStr(dr.Item(4))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(4) IsNot Nothing Then GetTitle = CStr(drLocal.Item(4))
                End While
                atsLocal.Close()
            End If
        Catch
            GetTitle = ""
        End Try
    End Function

    Public Function GetTraceID(Title As String, TestID As Long) As String
        Dim SQLStr As String
        Dim Count As Integer = 0
        Try

            GetTraceID = 0
            SQLStr = "SELECT * from Trace where JobNumber = '" & frmAUTOTEST.cmbJob.Text & "' And Title = '" & Title & "' And  Workstation = '" & GetComputerName() & "' And TestID = " & TestID & " And SerialNumber = '" & SerialNumber & "'"
            If SQL.CheckforRow(SQLStr, "NetworkTraceData") = 0 Then
                SQLStr = "Insert Into Trace (JobNumber, Title, Workstation,TestID,SerialNumber) values ('" & frmAUTOTEST.cmbJob.Text & "','" & Title & "','" & GetComputerName() & "'," & TestID & ",'" & SerialNumber & "')"
                SQL.ExecuteSQLCommand(SQLStr, "NetworkTraceData")
            End If
            SQLStr = "SELECT * from Trace where JobNumber = '" & frmAUTOTEST.cmbJob.Text & "' And Title = '" & Title & "' And  Workstation = '" & GetComputerName() & "' And TestID = " & TestID & " And SerialNumber = '" & SerialNumber & "'"
            If SQLAccess Then
                '****************************Get Trace Points********************************
                Dim ats1 As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd1 As SqlCommand = New SqlCommand(SQLStr, ats1)
                ats1.Open()

                Dim dr1 As SqlDataReader = cmd1.ExecuteReader()
                While Not dr1.Read = Nothing
                    If Not IsDBNull(dr1.GetValue(0)) Then GetTraceID = CType(dr1.Item(0), String)
                End While
                dr1.Close()

                '*******************************************************************************
            Else
                '****************************Get Trace Points********************************
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal1 As New OleDb.OleDbConnection
                atsLocal1.ConnectionString = strConnectionString
                Dim cmd1 As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal1)
                atsLocal1.Open()

                Dim drLocal1 As OleDb.OleDbDataReader = cmd1.ExecuteReader
                While Not drLocal1.Read = Nothing
                    If Not IsDBNull(drLocal1.GetValue(0)) Then GetTraceID = CType(drLocal1.Item(0), String)
                End While
                drLocal1.Close()
            End If
            '*******************************************************************************
        Catch ex As Exception
            MsgBox("Critical Error during GetTraceID! " & ex.Message, MsgBoxStyle.Critical, "Error")
            GetTraceID = 0
        End Try

    End Function
    Public Sub SaveTrace(Title As String, TestID As Long, TraceID As Long)
        Try
            Dim SQLStr As String = ""
            Dim XString As String = ""
            Dim yString As String = ""
            Dim Expression As String
            Dim Index As Integer

            'Populate Instrument Row
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                cmd.Connection = ats
                'Insert Initial Row
                Expression = " where ID = " & TraceID & " and Title = '" & Title & "'"
                cmd.CommandText = "UPDATE Trace SET  Workstation = '" & GetComputerName() & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET Points = '" & globals.Pts & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET JobNumber = '" & globals.Job & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET SerialNumber = '" & globals.SerialNumber & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET ActiveDate = '" & Now.Date.ToShortDateString & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET Temperature = '" & globals.Temperature & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET CalibrationDate = '" & globals.CalDate & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET InstrumentCalDue = '" & CalDue & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET ProgTitle  = 'IPP_CouplerTesting'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET ProgVersion = '" & globals.Version & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET XTitle = '" & globals.xTitle & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET  YTitle = '" & globals.yTitle & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET  SpecID = '" & SpecID & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET  artwork_rev = '" & ArtworkRevision & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET  Operator = '" & User & "'" & Expression
                cmd.ExecuteNonQuery()

                'Load String
                'note attempted to switch to a string trace, but SQL would not accept it.

                ' XString = StringifyTrace(XArray)
                '' yString = StringifyTrace(YArray)
                ' cmd.CommandText = "Insert Into TraceStr (TraceID,XData,YData) values (" & TraceID & "," & XString & "," & yString & ")"
                ' cmd.ExecuteNonQuery()

                'Load Trace Points
                'Note: switched to TracePoints2 after TraceID 171666
                For Index = 0 To XArray.Count - 1 Step 1
                    cmd.CommandText = "Insert Into TracePoints2 (TraceID, Idx, xData,YData) values (" & TraceID & "," & Index & "," & XArray(Index) & "," & YArray(Index) & ")"
                    cmd.ExecuteNonQuery()
                Next

                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                cmd.Connection = atsLocal
                'Insert Initial Row
                'SQLStr = "Insert Into Trace (Title,SerialNumber,JobNumber,ID) values ('" & globals.Title & " ', '" & globals.SerialNumber & " ','" & frmAUTOTEST.cmbJob.Text & "','" & globals.TraceID & "')"


                Expression = " where ID = " & TraceID & " and Title = '" & Title & "' And TestID = " & TestID

                cmd.CommandText = "UPDATE Trace SET  Workstation = '" & GetComputerName() & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET Points = '" & Pts & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET JobNumber = '" & globals.Job & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET SerialNumber = '" & globals.SerialNumber & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET ActiveDate = '" & Now.Date.ToShortDateString & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET Temperature = '" & globals.Temperature & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET CalibrationDate = '" & globals.CalDate & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET InstrumentCalDue = '" & globals.CalDue & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET ProgTitle  = 'IPP_CouplerTesting'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET ProgVer = '" & globals.Version & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET XTitle = '" & globals.xTitle & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET  YTitle = '" & globals.yTitle & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET  SpecID = " & SpecID & "" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET  artwork_rev = '" & ArtworkRevision & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace SET  Operator = '" & User & "'" & Expression
                cmd.ExecuteNonQuery()

                'Load String
                'note attempted to switch to a string trace, but SQL would not accept it.

                ' XString = StringifyTrace(XArray)
                '' yString = StringifyTrace(YArray)
                ' cmd.CommandText = "Insert Into TraceStr (TraceID,XData,YData) values (" & TraceID & "," & XString & "," & yString & ")"
                ' cmd.ExecuteNonQuery()

                'Load Trace Points
                'Note: switched to TracePoints2 after TraceID 171666
                For Index = 0 To XArray.Count - 1 Step 1
                    cmd.CommandText = "Insert Into TracePoints2 (TraceID, Idx, xData,YData) values (" & TraceID & "," & Index & "," & XArray(Index) & "," & YArray(Index) & ")"
                    cmd.ExecuteNonQuery()
                Next

                atsLocal.Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Error while inserting record on Trace.." & ex.Message, "Save Trace Data")
        End Try
    End Sub
    Public Sub SaveTraceImage(Title As String, TestID As Long, TraceID As Long)
        Try
            Dim SQLStr As String = ""
            Dim Expression As String
            Dim Index As Integer

            'Populate Instrument Row
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                cmd.Connection = ats
                'Insert Initial Row
                'SQLStr = "Insert Into Trace (Title,SerialNumber,JobNumber,ID) values ('" & globals.Title & " ', '" & globals.SerialNumber & " ','" & frmAUTOTEST.cmbJob.Text & "','" & globals.TraceID & "')"
                Expression = " where ID = " & TraceID & " and Title = '" & Title & "' And TestID = " & TestID

                cmd.CommandText = "UPDATE TraceImage SET  Workstation = '" & GetComputerName() & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET Points = '" & globals.Pts & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET JobNumber = '" & globals.Job & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET SerialNumber = '" & globals.SerialNumber & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET ActiveDate = '" & Now.Date.ToShortDateString & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET Temperature = '" & globals.Temperature & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET CalibrationDate = '" & globals.CalDate & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET InstrumentCalDue = '" & CalDue & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET ProgTitle  = 'IPP_CouplerTesting'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET ProgVersion = '" & globals.Version & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET XTitle = '" & globals.xTitle & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET  YTitle = '" & globals.yTitle & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET  SpecID = '" & SpecID & "'" & Expression
                cmd.ExecuteNonQuery()
                'Load Trace Points

                For Index = 0 To XArray.Count - 1 Step 1
                    cmd.CommandText = "Insert Into TraceImagePoints (TraceID, Idx, xData,YData) values (" & TraceID & "," & Index & "," & XArray(Index) & "," & YArray(Index) & ")"
                    cmd.ExecuteNonQuery()
                Next

                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                cmd.Connection = atsLocal
                'Insert Initial Row
                'SQLStr = "Insert Into Trace (Title,SerialNumber,JobNumber,ID) values ('" & globals.Title & " ', '" & globals.SerialNumber & " ','" & frmAUTOTEST.cmbJob.Text & "','" & globals.TraceID & "')"


                Expression = " where ID = " & TraceID & " and Title = '" & Title & "' And TestID = " & TestID

                cmd.CommandText = "UPDATE TraceImage SET  Workstation = '" & GetComputerName() & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET Points = '" & Pts & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET JobNumber = '" & globals.Job & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET SerialNumber = '" & globals.SerialNumber & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET ActiveDate = '" & Now.Date.ToShortDateString & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET Temperature = '" & globals.Temperature & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET CalibrationDate = '" & globals.CalDate & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET InstrumentCalDue = '" & globals.CalDue & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET ProgTitle  = 'IPP_CouplerTesting'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET ProgVer = '" & globals.Version & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET XTitle = '" & globals.xTitle & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET  YTitle = '" & globals.yTitle & "'" & Expression
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE TraceImage SET  SpecID = " & SpecID & "" & Expression
                cmd.ExecuteNonQuery()


                'Load Trace Points

                For Index = 0 To XArray.Count - 1 Step 1
                    cmd.CommandText = "Insert Into TraceImagePoints (TraceID, Idx, xData,YData) values (" & TraceID & "," & Index & "," & XArray(Index) & "," & YArray(Index) & ")"
                    cmd.ExecuteNonQuery()
                Next

                atsLocal.Close()
            End If
        Catch ex As Exception
            MessageBox.Show("Error while inserting record on Trace.." & ex.Message, "Save Trace Data")
        End Try
    End Sub


    Public Sub GetTrace(ByVal Title As String, Optional points As Integer = 201)
        Try
            Dim SQLStr As String
            Dim Count As Integer = 0

            'SEARCH BY BAND AND UUTType
            SQLStr = "Select * from Trace where Title = '" & Title & "' and TestID =  '" & TestID & "'"
            '*******************************************************************************
            'Load the Conguration Data
            '*******************************************************************************
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                globals.Notes = Nothing
                While Not dr.Read = Nothing
                    globals.TestID = CType(dr.Item(1), String)
                    globals.JobNumber = CType(dr.Item(3), String)
                    globals.Title = CType(dr.Item(4), String)
                    globals.SerialNumber = CType(dr.Item(5), String)
                    globals.WorkStation = CType(dr.Item(6), String)
                    globals.Pts = CType(dr.Item(7), String)
                    globals.ActiveDate = CType(dr.Item(8), String)
                    globals.CalDate = CType(dr.Item(11), String)
                    globals.CalDue = CType(dr.Item(12), String)
                    globals.ProgTitle = CType(dr.Item(13), String)
                    globals.ProgVer = CType(dr.Item(14), String)
                    globals.xTitle = CType(dr.Item(15), String)
                    globals.yTitle = CType(dr.Item(16), String)
                    If Not IsDBNull(dr.GetValue(17)) Then globals.Notes = CType(dr.Item(17), String)
                End While
                ats.Close()

            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                globals.Notes = Nothing
                While Not drLocal.Read = Nothing
                    globals.TestID = CType(drLocal.Item(1), String)
                    globals.JobNumber = CType(drLocal.Item(3), String)
                    globals.Title = CType(drLocal.Item(4), String)
                    globals.SerialNumber = CType(drLocal.Item(5), String)
                    globals.WorkStation = CType(drLocal.Item(6), String)
                    globals.Pts = CType(drLocal.Item(7), String)
                    globals.ActiveDate = CType(drLocal.Item(8), String)
                    globals.CalDate = CType(drLocal.Item(11), String)
                    globals.CalDue = CType(drLocal.Item(12), String)
                    globals.ProgTitle = CType(drLocal.Item(13), String)
                    globals.ProgVer = CType(drLocal.Item(14), String)
                    globals.xTitle = CType(drLocal.Item(15), String)
                    globals.yTitle = CType(drLocal.Item(16), String)
                    If Not IsDBNull(drLocal.GetValue(17)) Then globals.Notes = CType(drLocal.Item(17), String)
                End While
                atsLocal.Close()
            End If
            If TraceID > 171666 Then
                GetTracePoints2(TraceID, points)
            Else
                GetTracePoints(TraceID, points)
            End If


        Catch
        End Try
    End Sub

    Public Function GetTraceString(TraceID As Integer, Optional points As Integer = 201) As Boolean
        Dim SQLStr As String
        Dim Count As Integer = 0
        Try
            GetTraceString = False
            If SQLAccess Then
                '****************************Get Trace Points********************************
                SQLStr = "Select * from TracStr where TraceID = " & TraceID
                Dim ats1 As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd1 As SqlCommand = New SqlCommand(SQLStr, ats1)
                ats1.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr1 As SqlDataReader = cmd1.ExecuteReader()
                While Not dr1.Read = Nothing
                    If Not IsDBNull(dr1.GetValue(2)) Then XArray = ExpandTrace(CType(dr1.Item(2), String))
                    If Not IsDBNull(dr1.GetValue(3)) Then YArray = ExpandTrace(CType(dr1.Item(3), String))
                    GetTraceString = True
                    Count = Count + 1
                    If Count > points Then Exit While
                End While

                '*******************************************************************************
            Else
                '****************************Get Trace Points********************************
                SQLStr = "Select * from TraceStr where TraceID = " & TraceID
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal1 As New OleDb.OleDbConnection
                atsLocal1.ConnectionString = strConnectionString
                Dim cmd1 As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal1)
                atsLocal1.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal1 As OleDb.OleDbDataReader = cmd1.ExecuteReader
                While Not drLocal1.Read = Nothing
                    If Not IsDBNull(drLocal1.GetValue(2)) Then XArray = ExpandTrace(CType(drLocal1.Item(2), String))
                    If Not IsDBNull(drLocal1.GetValue(3)) Then YArray = ExpandTrace(CType(drLocal1.Item(3), String))
                    GetTraceString = True
                    Count = Count + 1
                    If Count > points Then Exit While
                End While
            End If
            '*******************************************************************************
        Catch
            GetTraceString = False
        End Try
    End Function
    Public Sub GetTracePoints(TraceID As Integer, Optional points As Integer = 201)
        Dim SQLStr As String
        Dim Count As Integer = 0
        Try

            If SQLAccess Then
                '****************************Get Trace Points********************************
                SQLStr = "Select * from TracePoints where TraceID = " & TraceID
                Dim ats1 As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd1 As SqlCommand = New SqlCommand(SQLStr, ats1)
                ats1.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr1 As SqlDataReader = cmd1.ExecuteReader()
                While Not dr1.Read = Nothing
                    ReDim Preserve XArray(Count)
                    ReDim Preserve YArray(Count)
                    If Not IsDBNull(dr1.GetValue(3)) Then XArray(Count) = CType(dr1.Item(3), String)
                    If Not IsDBNull(dr1.GetValue(4)) Then YArray(Count) = CType(dr1.Item(4), String)
                    Count = Count + 1
                    If Count > points Then Exit While
                End While

                '*******************************************************************************
            Else
                '****************************Get Trace Points********************************
                SQLStr = "Select * from TracePoints where TraceID = " & TraceID
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal1 As New OleDb.OleDbConnection
                atsLocal1.ConnectionString = strConnectionString
                Dim cmd1 As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal1)
                atsLocal1.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal1 As OleDb.OleDbDataReader = cmd1.ExecuteReader
                While Not drLocal1.Read = Nothing
                    ReDim Preserve XArray(Count)
                    ReDim Preserve YArray(Count)
                    If Not IsDBNull(drLocal1.GetValue(3)) Then XArray(Count) = CType(drLocal1.Item(3), String)
                    If Not IsDBNull(drLocal1.GetValue(4)) Then YArray(Count) = CType(drLocal1.Item(4), String)
                    Count = Count + 1
                    If Count > points Then Exit While
                End While
            End If
            '*******************************************************************************
        Catch
        End Try

    End Sub
    Public Sub GetTracePoints2(TraceID As Integer, Optional points As Integer = 201)
        Dim SQLStr As String
        Dim Count As Integer = 0
        Try

            If SQLAccess Then
                '****************************Get Trace Points********************************
                SQLStr = "Select * from TracePoints2 where TraceID = " & TraceID
                Dim ats1 As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd1 As SqlCommand = New SqlCommand(SQLStr, ats1)
                ats1.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr1 As SqlDataReader = cmd1.ExecuteReader()
                While Not dr1.Read = Nothing
                    ReDim Preserve XArray(Count)
                    ReDim Preserve YArray(Count)
                    If Not IsDBNull(dr1.GetValue(3)) Then XArray(Count) = CType(dr1.Item(3), String)
                    If Not IsDBNull(dr1.GetValue(4)) Then YArray(Count) = CType(dr1.Item(4), String)
                    Count = Count + 1
                    If Count > points Then Exit While
                End While

                '*******************************************************************************
            Else
                '****************************Get Trace Points********************************
                SQLStr = "Select * from TracePoints2 where TraceID = " & TraceID
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal1 As New OleDb.OleDbConnection
                atsLocal1.ConnectionString = strConnectionString
                Dim cmd1 As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal1)
                atsLocal1.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal1 As OleDb.OleDbDataReader = cmd1.ExecuteReader
                While Not drLocal1.Read = Nothing
                    ReDim Preserve XArray(Count)
                    ReDim Preserve YArray(Count)
                    If Not IsDBNull(drLocal1.GetValue(3)) Then XArray(Count) = CType(drLocal1.Item(3), String)
                    If Not IsDBNull(drLocal1.GetValue(4)) Then YArray(Count) = CType(drLocal1.Item(4), String)
                    Count = Count + 1
                    If Count > points Then Exit While
                End While
            End If
            '*******************************************************************************
        Catch
        End Try

    End Sub
    Public Sub GetTraceImage(ByVal Title As String)
        Try
            Dim SQLStr As String
            Dim Count As Integer = 0

            'SEARCH BY BAND AND UUTType
            SQLStr = "Select * from TraceImage where Title = '" & Title & "' and TestID =  '" & TestID & "'"
            '*******************************************************************************
            'Load the Conguration Data
            '*******************************************************************************
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    globals.TestID = CType(dr.Item(1), String)
                    globals.JobNumber = CType(dr.Item(3), String)
                    globals.Title = CType(dr.Item(4), String)
                    globals.SerialNumber = CType(dr.Item(5), String)
                    globals.WorkStation = CType(dr.Item(6), String)
                    globals.Pts = CType(dr.Item(7), String)
                    globals.ActiveDate = CType(dr.Item(8), String)
                    globals.CalDate = CType(dr.Item(11), String)
                    globals.CalDue = CType(dr.Item(12), String)
                    globals.ProgTitle = CType(dr.Item(13), String)
                    globals.ProgVer = CType(dr.Item(14), String)
                    globals.xTitle = CType(dr.Item(15), String)
                    globals.yTitle = CType(dr.Item(16), String)
                End While
                ats.Close()

            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    globals.TestID = CType(drLocal.Item(1), String)
                    globals.JobNumber = CType(drLocal.Item(3), String)
                    globals.Title = CType(drLocal.Item(4), String)
                    globals.SerialNumber = CType(drLocal.Item(5), String)
                    globals.WorkStation = CType(drLocal.Item(6), String)
                    globals.Pts = CType(drLocal.Item(7), String)
                    globals.ActiveDate = CType(drLocal.Item(8), String)
                    globals.CalDate = CType(drLocal.Item(11), String)
                    globals.CalDue = CType(drLocal.Item(12), String)
                    globals.ProgTitle = CType(drLocal.Item(13), String)
                    globals.ProgVer = CType(drLocal.Item(14), String)
                    globals.xTitle = CType(drLocal.Item(15), String)
                    globals.yTitle = CType(drLocal.Item(16), String)
                End While
                atsLocal.Close()
            End If
            GetTraceTraceImagePoints(TraceID)
        Catch
        End Try
    End Sub

    Public Sub GetTraceTraceImagePoints(TraceID As Integer)
        Dim SQLStr As String
        Dim Count As Integer = 0
        Try

            If SQLAccess Then
                '****************************Get Trace Points********************************
                SQLStr = "Select * from TracePoints where TraceID = " & TraceID
                Dim ats1 As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd1 As SqlCommand = New SqlCommand(SQLStr, ats1)
                ats1.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr1 As SqlDataReader = cmd1.ExecuteReader()
                While Not dr1.Read = Nothing
                    ReDim Preserve XArray(Count)
                    ReDim Preserve YArray(Count)
                    If Not IsDBNull(dr1.GetValue(3)) Then XArray(Count) = CType(dr1.Item(3), String)
                    If Not IsDBNull(dr1.GetValue(4)) Then YArray(Count) = CType(dr1.Item(4), String)
                    Count = Count + 1
                End While

                '*******************************************************************************
            Else
                '****************************Get Trace Points********************************
                SQLStr = "Select * from TracePoints where TraceID = " & TraceID
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal1 As New OleDb.OleDbConnection
                atsLocal1.ConnectionString = strConnectionString
                Dim cmd1 As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal1)
                atsLocal1.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal1 As OleDb.OleDbDataReader = cmd1.ExecuteReader
                While Not drLocal1.Read = Nothing
                    ReDim Preserve XArray(Count)
                    ReDim Preserve YArray(Count)
                    If Not IsDBNull(drLocal1.GetValue(3)) Then XArray(Count) = CType(drLocal1.Item(3), String)
                    If Not IsDBNull(drLocal1.GetValue(4)) Then YArray(Count) = CType(drLocal1.Item(4), String)
                    Count = Count + 1
                End While
            End If
            '*******************************************************************************
        Catch
        End Try

    End Sub

    Public Function GetPartNumber() As String
        Try
            Dim CountRow As Integer = 0
            Dim SQLstr As String
            SQLstr = "Select * from Specifications where JobNumber ='" & Job & "'"

            GetPartNumber = ""
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(3) IsNot Nothing Then GetPartNumber = CStr(dr.Item(3))
                End While
                ats.Close()
            Else
                System.Threading.Thread.Sleep(100)
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(3) IsNot Nothing Then GetPartNumber = CStr(drLocal.Item(3))
                End While
                atsLocal.Close()
            End If
        Catch
            GetPartNumber = ""
        End Try
    End Function
    Public Function GetTestStatus() As String
        Try
            Dim SQLstr As String
            Dim CountRow As Integer = 0
            Dim station As String
            Dim user As String
            Dim complete As String
            Dim total As String
            Dim status As String
            Dim thisJob As String

            SQLstr = "SELECT * from ReportQueue where JobNumber = '" & Job & "'"

            GetTestStatus = ""
            Dim x As Integer = 0
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(3) IsNot Nothing Then status = CStr(dr.Item(3))
                    If dr.Item(4) IsNot Nothing Then thisJob = CStr(dr.Item(4))
                    If dr.Item(5) IsNot Nothing Then station = CStr(dr.Item(5))
                    If dr.Item(7) IsNot Nothing Then user = CStr(dr.Item(7))
                    If dr.Item(10) IsNot Nothing Then complete = CStr(dr.Item(10))
                    If dr.Item(11) IsNot Nothing Then total = CStr(dr.Item(11))
                    ReDim Preserve ReportStatus(x)
                    ReportStatus(x) = status
                    ReDim Preserve SavedWorkStation(x)
                    SavedWorkStation(x) = station
                    ReDim Preserve SavedUser(x)
                    SavedUser(x) = user
                    ReDim Preserve SavedComplete(x)
                    SavedComplete(x) = complete
                    ReDim Preserve SavedTotal(x)
                    SavedTotal(x) = total
                    ReDim Preserve SavedJob(x)
                    SavedJob(x) = thisJob
                    x += 1

                End While
                ats.Close()
            Else
                System.Threading.Thread.Sleep(100)
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkSpecs")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.001)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(3) IsNot Nothing Then status = CStr(drLocal.Item(3))
                    If drLocal.Item(4) IsNot Nothing Then thisJob = CStr(drLocal.Item(4))
                    If drLocal.Item(5) IsNot Nothing Then station = CStr(drLocal.Item(5))
                    If drLocal.Item(7) IsNot Nothing Then user = CStr(drLocal.Item(7))
                    If drLocal.Item(10) IsNot Nothing Then complete = CStr(drLocal.Item(10))
                    If drLocal.Item(11) IsNot Nothing Then total = CStr(drLocal.Item(11))
                    ReDim Preserve ReportStatus(x)
                    ReportStatus(x) = status
                    ReDim Preserve SavedWorkStation(x)
                    SavedWorkStation(x) = station
                    ReDim Preserve SavedUser(x)
                    SavedUser(x) = user
                    ReDim Preserve SavedComplete(x)
                    SavedComplete(x) = complete
                    ReDim Preserve SavedTotal(x)
                    SavedTotal(x) = total
                    x += 1
                End While
                atsLocal.Close()
            End If
            If ReportStatus.Length = 0 Then
                GetTestStatus = "None"
            ElseIf ReportStatus.Length = 1 Then
                GetTestStatus = "One"
            ElseIf ReportStatus.Length > 1 Then
                GetTestStatus = "Multiple"
            End If
        Catch
            GetTestStatus = ""
        End Try
    End Function





    Public Sub DeleteOperator(Value As String)
        Dim SQLStr As String
        Try
            SQLStr = "DELETE from Effeciency where Operator = '" & Value & "'"
            SQL.ExecuteSQLCommand(SQLStr, "NetworkSpecs")


        Catch ex As Exception

        End Try


    End Sub

End Module