Imports System.Data
Imports System.Data.SqlClient

Public Class Trace
    Public Function OpenTrace() As Integer
        Dim SQLstr As String
        Dim Title As String
        Try
            OpenTrace = 0
            Title = " New Title"
            SQLstr = "SELECT * from Trace where JobNumber = '" & frmAUTOTEST.cmbJob.Text & "' And Title = '" & Title & "'"
            If SQL.CheckforRow(SQLstr, "LocalTraceData") = 0 Then
                SQLstr = "Insert Into Trace (JobNumber, Title) values ('" & frmAUTOTEST.cmbJob.Text & "','" & Title & "')"
                SQL.ExecuteSQLCommand(SQLstr, "LocalTraceData")
            End If

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand("SELECT * from Trace where JobNumber = '" & frmAUTOTEST.cmbJob.Text & "' And Title = '" & Title & "'", ats)
                ats.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    OpenTrace = CInt(dr.Item(0))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * from Trace where JobNumber = '" & frmAUTOTEST.cmbJob.Text & "' And Title = '" & Title & "'", atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    OpenTrace = drLocal.Item(0)
                End While
                atsLocal.Close()
            End If
        Catch ex As Exception
            OpenTrace = 0
        End Try
    End Function


    Public Sub AddPoint(TraceID As Long, Index As Long, XData As Double, YData As Double)
        Dim SQLstr As String

        Try
            SQLstr = "Insert Into TracePoints (TraceID,XData,YData,Idx) values ('" & TraceID & "','" & XData & "','" & YData & "','" & Index & "')"
            SQL.ExecuteSQLCommand(SQLstr, "LocalTraceData")
        Catch ex As Exception
        End Try
    End Sub


    Public Sub AddCoeffCalChar(Name As String, Coeff As String)
        Dim SQLstr As String
        Try
            SQLstr = "SELECT * from  CoeffCal where StringName =  '" & Name & "'"
            If SQL.CheckforRow(SQLstr, "LocalTraceData") = 0 Then
                SQLstr = "Insert Into CoeffCal (StringName) values ('" & Name & "')"
                SQL.ExecuteSQLCommand(SQLstr, "LocalTraceData")
            End If
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE Trace CoeffCal Set StringName = '" & Name & "' where StringName =  '" & Name & "'"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace CoeffCal Set String = '" & Coeff & "' where StringName =  '" & Name & "'"
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.Connection = atsLocal
                cmd.CommandText = "UPDATE Trace CoeffCal Set StringName = '" & Name & "' where StringName =  '" & Name & "'"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace CoeffCal Set String = '" & Coeff & "' where StringName =  '" & Name & "'"
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If
        Catch ex As Exception

        End Try
    End Sub


    Public Sub SaveTrace(TraceID As Long, Optional ShowPlot As Boolean = False)

        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE Trace Set ProgVersion = 'Version " & CType(System.Windows.Forms.Application.ProductVersion, String) & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace Set ProgTitle = '" & CType(System.Windows.Forms.Application.ProductName, String) & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace Set ActiveDate = '" & Now & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace Set InstrumentCalDue = '" & GetInstrCalDue() & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.CommandText = "UPDATE Trace Set ProgVer = 'Version " & CType(System.Windows.Forms.Application.ProductVersion, String) & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace Set ProgTitle = '" & CType(System.Windows.Forms.Application.ProductName, String) & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace Set ActiveDate = '" & Now & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                cmd.CommandText = "UPDATE Trace Set InstrumentCalDue = '" & GetInstrCalDue() & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If
        Catch ex As Exception
        End Try
    End Sub


    Public Sub SerialNumber(TraceID As Long, TraceSerialNumber As String)

        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE Trace Set SerialNumber = '" & TraceSerialNumber & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.CommandText = "UPDATE Trace Set SerialNumber = '" & TraceSerialNumber & "' where ID =  " & TraceID & """"
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If

        Catch ex As Exception

        End Try
    End Sub


    Public Sub Title(TraceID As Long, TraceTitle As String)

        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE Trace Set Title = '" & TraceTitle & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.CommandText = "UPDATE Trace Set Title = '" & TraceTitle & "' where ID =  " & TraceID & """"
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If

        Catch ex As Exception
        End Try
    End Sub


    Public Sub CalDate(TraceID As Long, Calibration As Date)

        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE Trace Set CalibrationDate = '" & Calibration & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.CommandText = "UPDATE Trace Set CalibrationDate = '" & Calibration & "' where ID =  " & TraceID & """"
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If

        Catch ex As Exception
        End Try
    End Sub

    Public Sub Notes(TraceID As Long, TestNotes As String)

        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE Trace Set Notes = '" & TestNotes & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.CommandText = "UPDATE Trace Set Notes = '" & TestNotes & "' where ID =  " & TraceID & """"
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If

        Catch ex As Exception
        End Try

    End Sub

    Public Sub Points(TraceID As Long, TracePoints As Integer)

        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE Trace Set Points = '" & TracePoints & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.CommandText = "UPDATE Trace Set Points = '" & TracePoints & "' where ID =  " & TraceID & """"
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If

        Catch ex As Exception
        End Try
    End Sub

    Public Sub SpecID(TraceID As Long, SpecificationID As Long)

        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE Trace Set SpecID = '" & SpecificationID & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.CommandText = "UPDATE Trace Set SpecID = '" & SpecificationID & "' where ID =  " & TraceID & """"
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If

        Catch ex As Exception
        End Try
    End Sub

    Public Sub TestID(TraceID As Long, TestIDs As Long)
        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE Trace Set TestID = '" & TestIDs & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.CommandText = "UPDATE Trace Set TestID = '" & TestIDs & "' where ID =  " & TraceID & """"
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If

        Catch ex As Exception
        End Try

    End Sub

    Public Sub Workstation(TraceID As Long, ComputerName As String)
        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand()
                ats.Open()
                cmd.Connection = ats
                cmd.CommandText = "UPDATE Trace Set Workstation = '" & ComputerName & "' where ID =  " & TraceID & ""
                cmd.ExecuteNonQuery()
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                atsLocal.Open()
                cmd.CommandText = "UPDATE Trace Set Workstation = '" & ComputerName & "' where ID =  " & TraceID & """"
                cmd.ExecuteNonQuery()
                atsLocal.Close()
            End If

        Catch ex As Exception
        End Try

    End Sub

    


    Public Sub MinMax(TraceID As Long)
        Dim i As Integer = 0
        Try
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand("SELECT * from TracePoints where TraceID =  " & TraceID & "", ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                i = 0
                While Not dr.Read = Nothing
                    If i = 0 Then
                        retMax = CDbl(dr.Item(4))
                        retMin = CDbl(dr.Item(4))
                    Else
                        If CDbl(dr.Item(0)) > retMax Then
                            retMax = CDbl(dr.Item(4))
                            retMaxX = i
                        End If
                        If CDbl(dr.Item(0)) < retMin Then
                            retMin = CDbl(dr.Item(4))
                            retMinX = i
                        End If
                    End If
                    i = i + 1
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * from TracePoints where TraceID =  " & TraceID & "", atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                i = 0
                While Not drLocal.Read = Nothing
                    If i = 0 Then
                        retMax = CDbl(drLocal.Item(4))
                        retMin = CDbl(drLocal.Item(4))
                    Else
                        If CDbl(drLocal.Item(0)) > retMax Then
                            retMax = CDbl(drLocal.Item(4))
                            retMaxX = i
                        End If
                        If CDbl(drLocal.Item(0)) < retMin Then
                            retMin = CDbl(drLocal.Item(4))
                            retMinX = i
                        End If
                    End If
                    i = i + 1
                End While
                atsLocal.Close()
            End If
        Catch
        End Try
    End Sub


    Public Function GetYData(TraceID As Long, Index As Integer) As Double
        Try
            GetYData = 0
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand("SELECT * from TracePoints where TraceID =  " & TraceID & " and Idx = " & Index & "", ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.1)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    GetYData = CDbl(dr.Item(4))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * from TracePoints where TraceID =  " & TraceID & " and Idx = " & Index & "", atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.1)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    GetYData = CDbl(drLocal.Item(4))
                End While
                atsLocal.Close()
            End If
        Catch
            GetYData = 0
        End Try
    End Function


    Public Function GetXData(TraceID As Long, Index As Integer) As Double
        Try
            GetXData = 0
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand("SELECT * from TracePoints where TraceID =  " & TraceID & " and Idx = " & Index & "", ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.01)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    GetXData = CDbl(dr.Item(3))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * from TracePoints where TraceID =  " & TraceID & " and Idx = " & Index & "", atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.01)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    GetXData = CDbl(drLocal.Item(3))
                End While
                atsLocal.Close()
            End If
        Catch
            GetXData = 0
        End Try
    End Function



    Public Function GetPoints(TraceID As Long) As Double

        Try
            GetPoints = 201
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand("SELECT * from Trace where ID =  " & TraceID & "", ats)
                ats.Open()
                System.Threading.Thread.Sleep(0.01)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    GetPoints = CDbl(dr.Item(7))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("LocalTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * from Trace where ID =  " & TraceID & "", atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(0.01)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    GetPoints = CDbl(drLocal.Item(7))
                End While
                atsLocal.Close()
            End If
        Catch
            GetPoints = 201
        End Try

    End Function


    Public Function GetInstrCalDue() As Date
        Try
            GetInstrCalDue = Date.Now
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand("SELECT * from InstCalDue where WorkStation =  '" & GetComputerName() & "'", ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    GetInstrCalDue = dr.Item(2)
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * from InstCalDue where WorkStation =  '" & GetComputerName() & "'", atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    GetInstrCalDue = drLocal.Item(2)
                End While
                atsLocal.Close()
            End If
        Catch
            GetInstrCalDue = Date.Now
        End Try
        CalDue = GetInstrCalDue.ToString
        Exit Function
    End Function

    Public Sub GetCalCoeffs()
        Dim SQLstr As String
        Dim Title As String
        Dim Channel, SwPos As Integer
        Try

            CalCoeff1 = ""
            CalCoeff2 = ""
            CalCoeff3 = ""
            If SQLAccess Then
                'SwPos = 1
                Channel = 2
                SwPos = 1
                Title = "CalCoeff CH" & Channel & " SwPos: " & SwPos
                SQLstr = "SELECT * from CoeffCal where StringName =  '" & Title & "'"
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    CalCoeff1 = CalCoeff1 & dr.Item(12)
                End While
                ats.Close()

                'SwPos = 2
                Dim cmd1 As SqlCommand = New SqlCommand(SQLstr, ats)
                Channel = 2
                SwPos = 2
                Title = "CalCoeff CH" & Channel & " SwPos: " & SwPos
                SQLstr = "SELECT * from CoeffCal where StringName =  '" & Title & "'"
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr1 As SqlDataReader = cmd1.ExecuteReader()
                While Not dr1.Read = Nothing
                    CalCoeff2 = CalCoeff2 & dr1.Item(2)
                End While
                ats.Close()

                'SwPos = 3
                Dim cmd2 As SqlCommand = New SqlCommand(SQLstr, ats)
                Channel = 2
                SwPos = 3
                Title = "CalCoeff CH" & Channel & " SwPos: " & SwPos
                SQLstr = "SELECT * from CoeffCal where StringName =  '" & Title & "'"
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr2 As SqlDataReader = cmd2.ExecuteReader()
                While Not dr2.Read = Nothing
                    CalCoeff3 = CalCoeff3 & dr2.Item(2)
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkData")
                'SwPos = 1
                Channel = 2
                SwPos = 1
                Title = "CalCoeff CH" & Channel & " SwPos: " & SwPos
                SQLstr = "SELECT * from CoeffCal where StringName =  '" & Title & "'"
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    CalCoeff1 = CalCoeff1 & drLocal.Item(2)
                End While
                atsLocal.Close()

                'SwPos = 2
                Channel = 2
                SwPos = 2
                Title = "CalCoeff CH" & Channel & " SwPos: " & SwPos
                SQLstr = "SELECT * from CoeffCal where StringName =  '" & Title & "'"
                Dim cmd1 As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr1Local As OleDb.OleDbDataReader = cmd1.ExecuteReader
                While Not dr1Local.Read = Nothing
                    CalCoeff2 = CalCoeff2 & dr1Local.Item(2)
                End While
                atsLocal.Close()

                'SwPos = 3
                Channel = 2
                SwPos = 3
                Title = "CalCoeff CH" & Channel & " SwPos: " & SwPos
                SQLstr = "SELECT * from CoeffCal where StringName =  '" & Title & "'"
                Dim cmd2 As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr2Local As OleDb.OleDbDataReader = cmd2.ExecuteReader
                While Not dr1Local.Read = Nothing
                    CalCoeff3 = CalCoeff3 & dr1Local.Item(2)
                End While
                atsLocal.Close()

            End If
        Catch
        End Try
    End Sub

    Public Sub AddAllPoints(TraceID As Long, Pts As Long, XData() As Double, YData() As Double)
        Dim SQLstr As String
        Dim ATS As ADODB.Recordset
        Dim i As Integer
        On Error GoTo Trap

        i = 0
        Do While i < Pts
            SQLstr = "Insert Into TracePoints (TraceID,XData,YData,Idx) values ('" & TraceID & "','" & XData(i) & "','" & YData(i) & "','" & i & "')"
            SQL.ExecuteSQLCommand(SQLstr, "LocalTraceData")
            i = i + 1
        Loop

        Exit Sub
Trap:

    End Sub

End Class
