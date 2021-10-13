Imports System.Data
Imports System.Data.SqlClient

Module MSChart
    Dim ii As Double
    Dim iii As Integer
    Dim e As Double
    Dim r As Double
    Public chht As Long
    Public chwd As Long
    Dim chht1 As Long
    Dim chwd1 As Long
    Dim digChart As CommonDialog
    ' of the Form or Code Module.
    Public Function RedFromRGB(ByVal rgb As Long) _
       As Integer
        ' The ampersand after &HFF coerces the number as a
        ' long, preventing Visual Basic from evaluating the
        ' number as a negative value. The logical And is
        ' used to return bit values.
        RedFromRGB = &HFF& And rgb
    End Function

    Public Function GreenFromRGB(ByVal rgb As Long) _
       As Integer
        ' The result of the And operation is divided by
        ' 256, to return the value of the middle bytes.
        ' Note the use of the Integer divisor.
        GreenFromRGB = (&HFF00& And rgb) \ 256
    End Function

    Public Function BlueFromRGB(ByVal rgb As Long) _
       As Integer
        ' This function works like the GreenFromRGB above,
        ' except you don't need the ampersand. The
        ' number is already a long. The result divided by
        ' 65536 to obtain the highest bytes.
        BlueFromRGB = (&HFF0000 And rgb) \ 65536
    End Function

    Public Sub UpDateChartData(SpecType As String, Test As String, PassFail As String, Optional Retest As Boolean = False, Optional Undo As Boolean = False)
        Dim SQLstr As String
        Dim Pass As String = "Pass"
        Dim Fail As String = "Fail"
        Dim Expression As String
        Dim Table As String

        SQLstr = ""
        'Clear the Chart Data
        If InStr(SpecType, "90 DEGREE COUPLER") Or InStr(SpecType, "BALUN") Or InStr(SpecType, "COMBINER/DIVIDER") Then SQLstr = "select * from graphdb_3dB where Statename = '" & Test & "'"
        If InStr(SpecType, "SINGLE DIRECTIONAL COUPLER") Or InStr(SpecType, "DUAL DIRECTIONAL COUPLER") Or InStr(SpecType, "BI DIRECTIONAL COUPLER") Then SQLstr = "select * from graphdb_Dir where Statename = '" & Test & "'"
        Try
            'Get value from database
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If dr.Item(1) IsNot Nothing Then Pass = CStr(dr.Item(1))
                    If dr.Item(2) IsNot Nothing Then Pass = CStr(dr.Item(2))
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLstr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If drLocal.Item(1) IsNot Nothing Then Pass = CStr(drLocal.Item(1))
                    If drLocal.Item(2) IsNot Nothing Then Pass = CStr(drLocal.Item(2))
                End While
                atsLocal.Close()
            End If
            Table = "graphdb_3dB"
            'update database
            Expression = " where JobNumber = '" & Job & "' And SerialNumber = '" & SerialNumber & "'"
            If InStr(SpecType, "90 DEGREE COUPLER") Or InStr(SpecType, "BALUN") Or InStr(SpecType, "COMBINER/DIVIDER") Then Table = "graphdb_3dB"
            If InStr(SpecType, "SINGLE DIRECTIONAL COUPLER") Or InStr(SpecType, "DUAL DIRECTIONAL COUPLER") Or InStr(SpecType, "BI DIRECTIONAL COUPLER") Then Table = "graphdb_Dir"
            If SQLAccess Then
                Dim ats1 As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd1 As SqlCommand = New SqlCommand()
                ats1.Open()
                cmd1.Connection = ats1
                If PassFail = "Pass" And Not Undo And Not Retest Then
                    cmd1.CommandText = "UPDATE from " & Table & " Set Pass = '" & Pass + 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                End If
                If PassFail = "Pass" And Undo And Not Retest Then
                    cmd1.CommandText = "UPDATE from " & Table & " Set Pass = '" & Pass - 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                End If

                If PassFail = "Fail" And Not Undo And Not Retest Then
                    cmd1.CommandText = "UPDATE from " & Table & " Set Fail = '" & Fail + 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                End If
                If PassFail = "Fail" And Undo And Not Retest Then
                    cmd1.CommandText = "UPDATE from " & Table & " Set Fail = '" & Fail - 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                End If

                If PassFail = "Pass" And Retest Then
                    cmd1.CommandText = "UPDATE from " & Table & " Set Fail = '" & Fail - 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                    cmd1.CommandText = "UPDATE from " & Table & " Set Pass = '" & Pass + 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                End If
                ats1.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkData")
                Dim ats1Local As New OleDb.OleDbConnection
                ats1Local.ConnectionString = strConnectionString
                Dim cmd1 As OleDb.OleDbCommand = New OleDb.OleDbCommand()
                ats1Local.Open()
                cmd1.Connection = ats1Local
                If PassFail = "Pass" And Not Undo And Not Retest Then
                    cmd1.CommandText = "UPDATE from " & Table & " Set Pass = '" & Pass + 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                End If
                If PassFail = "Pass" And Undo And Not Retest Then
                    cmd1.CommandText = "UPDATE from " & Table & " Set Pass = '" & Pass - 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                End If

                If PassFail = "Fail" And Not Undo And Not Retest Then
                    cmd1.CommandText = "UPDATE from " & Table & " Set Fail = '" & Fail + 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                End If
                If PassFail = "Fail" And Undo And Not Retest Then
                    cmd1.CommandText = "UPDATE from " & Table & " Set Fail = '" & Fail - 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                End If

                If PassFail = "Pass" And Retest Then
                    cmd1.CommandText = "UPDATE from " & Table & " Set Fail = '" & Fail - 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                    cmd1.CommandText = "UPDATE from " & Table & " Set Pass = '" & Pass + 1 & "'" & Expression
                    cmd1.ExecuteNonQuery()
                End If
                ats1Local.Close()
            End If

        Catch ex As Exception
        End Try
    End Sub

    Public Sub ResetChartData(SpecType As String)
        Dim SQLstr As String = ""
        If SpecType = "" Then Exit Sub
        'Clear the Chart Data
        If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then SQLstr = "UPDATE from Graphdb_3dB Set "
        If SpecType = "90 DEGREE COUPLER SMD" Or SpecType = "BALUN SMD" Or SpecType = "COMBINER/DIVIDER SMD" Then SQLstr = "UPDATE from Graphdb_3dB Set "
        If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then SQLstr = "UPDATE from Graphdb_Dir Set "
        If SpecType = "SINGLE DIRECTIONAL COUPLER SMD" Or SpecType = "DUAL DIRECTIONAL COUPLER SMD" Or SpecType = "BI DIRECTIONAL COUPLER SMD" Then SQLstr = "UPDATE from Graphdb_Dir Set "
        If SpecType.Contains("COMBINER/DIVIDER") Or SpecType = "COMBINER/DIVIDER SMD" Then SQLstr = "UPDATE from Graphdb_3dB Set "
        If SQLAccess Then
            Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
            Dim cmd As SqlCommand = New SqlCommand()
            ats.Open()
            cmd.Connection = ats
            cmd.CommandText = SQLstr & "Pass = 0"
            cmd.ExecuteNonQuery()
            cmd.CommandText = SQLstr & "Fail = 0"
            cmd.ExecuteNonQuery()
            ats.Close()
        Else
            Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("LocalTraceData")
            Dim atsLocal As New OleDb.OleDbConnection
            atsLocal.ConnectionString = strConnectionString
            Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand()
            atsLocal.Open()
            cmd.Connection = atsLocal
            cmd.CommandText = SQLstr & "Pass = 0"
            cmd.ExecuteNonQuery()
            cmd.CommandText = SQLstr & "Fail = 0"
            cmd.ExecuteNonQuery()
            atsLocal.Close()
        End If

    End Sub

    Public Sub InitializeChart()
        'Dim Red, Green, Blue As Integer


        '    Set digChart = New CommonDialog
        '    With dlgChart ' CommonDialog object
        '      .CancelError = True
        '      .ShowColor
        '      Red = RedFromRGB(.Color)
        '      Green = GreenFromRGB(.Color)
        '      Blue = BlueFromRGB(.Color)
        '    End With

        '    frmAUTOTEST.MSChart1.ToDefaults
        '    frmAUTOTEST.MSChart1.ShowLegend = True
        '    frmAUTOTEST.MSChart1.RowCount = 6
        'frmAUTOTEST.MSChart1.Plot.DataSeriesInRow = False

        ' frmAUTOTEST.MSChart1.Plot.Wall.Brush.FillColor.Set 255, 255, 255
        'frmAUTOTEST.MSChart1.Plot.Backdrop.Frame.Style = VtFrameStyleThickOuter
    End Sub

    Public Function UpDateChart(SpecType As String)
        '    Dim SQLstr As String
        '    Dim ATS As ADODB.Recordset
        '    If SpecType = "" Then Exit Function
        '    If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then SQLstr = "select * from Graphdb_3dB"
        '    If SpecType = "90 DEGREE COUPLER SMD" Or SpecType = "BALUN SMD" Or SpecType = "COMBINER/DIVIDER SMD" Then SQLstr = "select * from Graphdb_3dB"
        '    If SpecType = "SINGLE DIRECTIONAL COUPLER" Or SpecType = "DUAL DIRECTIONAL COUPLER" Or SpecType = "BI DIRECTIONAL COUPLER" Then SQLstr = "select * from Graphdb_Dir"
        '    If SpecType = "SINGLE DIRECTIONAL COUPLER SMD" Or SpecType = "DUAL DIRECTIONAL COUPLER SMD" Or SpecType = "BI DIRECTIONAL COUPLER SMD" Then SQLstr = "select * from Graphdb_Dir"
        '    If SpecType.Contains("COMBINER/DIVIDER") Or SpecType = "COMBINER/DIVIDER SMD" Then SQLstr = "select * from Graphdb_3dB"
        '    ATS = New ADODB.Recordset
        '    ATS.Open(SQLstr, LocalDataConn, adOpenDynamic, adLockOptimistic)
        '    frmAUTOTEST.MSChart1.Enabled = True

        '    With frmAUTOTEST.MSChart1
        '        .DataSource = ATS
        '        .ShowLegend = True
        '    End With
        Return ""
    End Function
End Module
