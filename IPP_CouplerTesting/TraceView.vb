Imports System.Data
Imports System.Data.SqlClient
Imports ZedGraph



Public Class TraceView
    Private Points As Integer
    Private TraceID As Integer
    Private Trace1XArray(Points) As Double
    Private Trace1YArray(Points) As Double
    Private Trace2XArray(Points) As Double
    Private Trace2YArray(Points) As Double
    Private Trace3XArray(Points) As Double
    Private Trace3YArray(Points) As Double
    Private Trace4XArray(Points) As Double
    Private Trace4YArray(Points) As Double
    Private Trace1 As LineItem
    Private Trace2 As LineItem
    Private Trace3 As LineItem
    Private Trace4 As LineItem
    Private list1 As New PointPairList
    Private list2 As New PointPairList
    Private list3 As New PointPairList
    Private list4 As New PointPairList
    Private myPane As GraphPane
    Private SerialListLoad As Boolean = False
    Private TraceSN As String
    Private TraceJob As String
    Private TracePart As String
    Private TraceTitle As String
    Private TraceTestID As String

    Public Sub New()
        MyBase.New()
        Dim SQLStr As String

        Try
            InitializeComponent()
            SerialListLoad = False
            cmbPARM1.Text = ""
            txtPARM1.Visible = True
            TraceGrid.Visible = True
            SQLStr = "Select * from Trace"
            LoadTestLists()
            LoadTraceView(SQLStr)
            myPane.Chart.Fill.SecondaryValueGradientColor = Color.Black
            myPane.Title.Text = "System Noise Temperature"
            myPane.XAxis.Title.Text = "Frequency (MHz)"
            myPane.XAxis.MajorGrid.IsVisible = True
            myPane.XAxis.MajorGrid.DashOff = 0
            myPane.XAxis.MajorGrid.Color = Color.White
            myPane.YAxis.Title.Text = "dB"
            myPane.YAxis.MajorGrid.IsVisible = True
            myPane.YAxis.MajorGrid.DashOff = 0
            myPane.YAxis.MajorGrid.Color = Color.White
            myPane.Chart.Fill.Color = Color.Black
            myPane.Chart.Fill.Brush = Brushes.Black
            myPane.Fill.Color = Color.LightGray
            Me.Refresh()

        Catch
        End Try
    End Sub
    Private Sub LoadTraceView(SQLStr As String)
        Dim NameStr(18) As String
        Try
            TraceGrid.Visible = True
            ZG1.Visible = False
            TraceGrid.Rows.Clear()
            TraceGrid.Columns.Clear()
            NameStr(0) = "TraceID"
            NameStr(1) = "TestID"
            NameStr(2) = "SpecID"
            NameStr(3) = "JobNumber"
            NameStr(4) = "PartNumber"
            NameStr(5) = "Title"
            NameStr(6) = "SerialNumber"
            NameStr(7) = "WorkStation"
            NameStr(8) = "Points"
            NameStr(9) = "ActiveDate"
            NameStr(10) = "RFPower"
            NameStr(11) = "Temperature"
            NameStr(12) = "CalibrationDate"
            NameStr(13) = "InstrumentCalDue"
            NameStr(14) = "ProgTitle"
            NameStr(15) = "ProgVer"
            NameStr(16) = "XTitle"
            NameStr(17) = "YTitle"
            NameStr(18) = "Notes"
           

            For x = 0 To 18
                Dim col As New DataGridViewTextBoxColumn

                col.DataPropertyName = NameStr(x)
                col.HeaderText = NameStr(x)
                col.Name = NameStr(x)
                TraceGrid.Columns.Add(col)
            Next

            If CheckforRow(SQLStr, "NetworkTraceData") > 0 Then GoTo GotIt

GotIt:
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    Dim Count As Integer = TraceGrid.Rows.Add()

                    If Not IsDBNull(dr.Item(0)) Then TraceGrid.Rows.Item(Count).Cells(0).Value = CType(dr.Item(0), String)
                    If Not IsDBNull(dr.Item(1)) Then TraceGrid.Rows.Item(Count).Cells(1).Value = CType(dr.Item(1), String)
                    If Not IsDBNull(dr.Item(2)) Then TraceGrid.Rows.Item(Count).Cells(2).Value = CType(dr.Item(2), String)
                    If Not IsDBNull(dr.Item(3)) Then TraceGrid.Rows.Item(Count).Cells(3).Value = CType(dr.Item(3), String)
                    If Not IsDBNull(dr.Item(4)) Then TraceGrid.Rows.Item(Count).Cells(4).Value = CType(dr.Item(4), String)
                    If Not IsDBNull(dr.Item(5)) Then TraceGrid.Rows.Item(Count).Cells(5).Value = CType(dr.Item(5), String)
                    If Not IsDBNull(dr.Item(6)) Then TraceGrid.Rows.Item(Count).Cells(6).Value = CType(dr.Item(6), String)
                    If Not IsDBNull(dr.Item(7)) Then TraceGrid.Rows.Item(Count).Cells(7).Value = CType(dr.Item(7), String)
                    If Not IsDBNull(dr.Item(8)) Then TraceGrid.Rows.Item(Count).Cells(8).Value = CType(dr.Item(8), String)
                    If Not IsDBNull(dr.Item(9)) Then TraceGrid.Rows.Item(Count).Cells(9).Value = CType(dr.Item(9), String)
                    If Not IsDBNull(dr.Item(10)) Then TraceGrid.Rows.Item(Count).Cells(10).Value = CType(dr.Item(10), String)
                    If Not IsDBNull(dr.Item(11)) Then TraceGrid.Rows.Item(Count).Cells(11).Value = CType(dr.Item(11), String)
                    If Not IsDBNull(dr.Item(12)) Then TraceGrid.Rows.Item(Count).Cells(12).Value = CType(dr.Item(12), String)
                    If Not IsDBNull(dr.Item(13)) Then TraceGrid.Rows.Item(Count).Cells(13).Value = CType(dr.Item(13), String)
                    If Not IsDBNull(dr.Item(14)) Then TraceGrid.Rows.Item(Count).Cells(14).Value = CType(dr.Item(14), String)
                    If Not IsDBNull(dr.Item(15)) Then TraceGrid.Rows.Item(Count).Cells(15).Value = CType(dr.Item(15), String)
                    If Not IsDBNull(dr.Item(16)) Then TraceGrid.Rows.Item(Count).Cells(16).Value = CType(dr.Item(16), String)
                    If Not IsDBNull(dr.Item(17)) Then TraceGrid.Rows.Item(Count).Cells(17).Value = CType(dr.Item(17), String)
                    ' If Not IsDBNull(dr.Item(18)) Then TraceGrid.Rows.Item(Count).Cells(18).Value = CType(dr.Item(18), String)
                   

                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    Dim Count As Integer = TraceGrid.Rows.Add()

                    If Not IsDBNull(drLocal.Item(0)) Then TraceGrid.Rows.Item(Count).Cells(0).Value = CType(drLocal.Item(0), String)
                    If Not IsDBNull(drLocal.Item(1)) Then TraceGrid.Rows.Item(Count).Cells(1).Value = CType(drLocal.Item(1), String)
                    If Not IsDBNull(drLocal.Item(2)) Then TraceGrid.Rows.Item(Count).Cells(2).Value = CType(drLocal.Item(2), String)
                    If Not IsDBNull(drLocal.Item(3)) Then TraceGrid.Rows.Item(Count).Cells(3).Value = CType(drLocal.Item(3), String)
                    If Not IsDBNull(drLocal.Item(4)) Then TraceGrid.Rows.Item(Count).Cells(4).Value = CType(drLocal.Item(4), String)
                    If Not IsDBNull(drLocal.Item(5)) Then TraceGrid.Rows.Item(Count).Cells(5).Value = CType(drLocal.Item(5), String)
                    If Not IsDBNull(drLocal.Item(6)) Then TraceGrid.Rows.Item(Count).Cells(6).Value = CType(drLocal.Item(6), String)
                    If Not IsDBNull(drLocal.Item(7)) Then TraceGrid.Rows.Item(Count).Cells(7).Value = CType(drLocal.Item(7), String)
                    If Not IsDBNull(drLocal.Item(8)) Then TraceGrid.Rows.Item(Count).Cells(8).Value = CType(drLocal.Item(8), String)
                    If Not IsDBNull(drLocal.Item(9)) Then TraceGrid.Rows.Item(Count).Cells(9).Value = CType(drLocal.Item(9), String)
                    If Not IsDBNull(drLocal.Item(10)) Then TraceGrid.Rows.Item(Count).Cells(10).Value = CType(drLocal.Item(10), String)
                    If Not IsDBNull(drLocal.Item(11)) Then TraceGrid.Rows.Item(Count).Cells(11).Value = CType(drLocal.Item(11), String)
                    If Not IsDBNull(drLocal.Item(12)) Then TraceGrid.Rows.Item(Count).Cells(12).Value = CType(drLocal.Item(12), String)
                    If Not IsDBNull(drLocal.Item(13)) Then TraceGrid.Rows.Item(Count).Cells(13).Value = CType(drLocal.Item(13), String)
                    If Not IsDBNull(drLocal.Item(14)) Then TraceGrid.Rows.Item(Count).Cells(14).Value = CType(drLocal.Item(14), String)
                    If Not IsDBNull(drLocal.Item(15)) Then TraceGrid.Rows.Item(Count).Cells(15).Value = CType(drLocal.Item(15), String)
                    If Not IsDBNull(drLocal.Item(16)) Then TraceGrid.Rows.Item(Count).Cells(16).Value = CType(drLocal.Item(16), String)
                    If Not IsDBNull(drLocal.Item(17)) Then TraceGrid.Rows.Item(Count).Cells(16).Value = CType(drLocal.Item(17), String)
                    'If Not IsDBNull(drLocal.Item(18)) Then TraceGrid.Rows.Item(Count).Cells(18).Value = CType(drLocal.Item(18), String)
                  
                End While
                atsLocal.Close()
            End If
            '*******************************************************************************
        Catch
        End Try
    End Sub


    Private Function LoadTestLists() As Integer
        Try
            Dim cnt As Int32 = 0
            Dim SQLStr As String
            Me.Trace1Sel.Items.Clear()
            Me.Trace2Sel.Items.Clear()
            Me.Trace3Sel.Items.Clear()
            Me.Trace4Sel.Items.Clear()

            SQLStr = "Select DISTINCT Title from Trace"
            If cmbPARM1.Text = "TraceID" Then cmbPARM1.Text = "ID"
            If cmbPARM1.Text = "TraceID" Or cmbPARM1.Text = "TestID" Then
                If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = " & txtPARM1.Text
            Else
                If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
            End If
            If cmbPARM2.Text = "TraceID" Then cmbPARM2.Text = "ID"
            If cmbPARM2.Text = "SerialNumber" Then
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & cmbSerialNumber1.Text & "'"
            ElseIf cmbPARM2.Text = "TraceID" Or cmbPARM2.Text = "TestID" Then
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = " & txtPARM2.Text
            Else
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
            End If
            If cmbPARM3.Text = "TraceID" Then cmbPARM3.Text = "ID"
            If cmbPARM3.Text = "SerialNumber" Then
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & cmbSerialNumber2.Text & "'"
            ElseIf cmbPARM3.Text = "TraceID" Or cmbPARM3.Text = "TestID" Then
                If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = " & txtPARM3.Text
            Else
                If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"
            End If

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If Not IsDBNull(dr.GetValue(0)) Then
                        Me.Trace1Sel.Items.Add(dr.GetValue(0))
                        Me.Trace2Sel.Items.Add(dr.GetValue(0))
                        Me.Trace3Sel.Items.Add(dr.GetValue(0))
                        Me.Trace4Sel.Items.Add(dr.GetValue(0))
                    End If
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    Me.Trace1Sel.Items.Add(drLocal.GetValue(0))
                    Me.Trace2Sel.Items.Add(drLocal.GetValue(0))
                    Me.Trace3Sel.Items.Add(drLocal.GetValue(0))
                    Me.Trace4Sel.Items.Add(drLocal.GetValue(0))
                End While
                atsLocal.Close()

            End If
        Catch
            Return 0
        End Try
        Return 1
    End Function

    Private Function LoadSNLists() As Integer
        Try
            Dim cnt As Int32 = 0
            Dim SQLStr As String
            Me.cmbSerialNumber1.Items.Clear()
            Me.cmbSerialNumber2.Items.Clear()
            

            SQLStr = "Select DISTINCT SerialNumber from Trace"
            If cmbPARM1.Text = "TraceID" Then cmbPARM1.Text = "ID"
            If cmbPARM1.Text = "TraceID" Or cmbPARM1.Text = "TestID" Then
                If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = " & txtPARM1.Text
            Else
                If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
            End If
            If cmbPARM2.Text = "TraceID" Then cmbPARM2.Text = "ID"
            If cmbPARM2.Text = "SerialNumber" Then
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & cmbSerialNumber1.Text & "'"
            ElseIf cmbPARM2.Text = "TraceID" Or cmbPARM2.Text = "TestID" Then
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = " & txtPARM2.Text
            Else
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
            End If
            If cmbPARM3.Text = "TraceID" Then cmbPARM3.Text = "ID"
            If cmbPARM3.Text = "SerialNumber" Then
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & cmbSerialNumber2.Text & "'"
            ElseIf cmbPARM3.Text = "TraceID" Or cmbPARM3.Text = "TestID" Then
                If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = " & txtPARM3.Text
            Else
                If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"
            End If

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If Not IsDBNull(dr.GetValue(0)) Then
                        Me.cmbSerialNumber1.Items.Add(dr.GetValue(0))
                        Me.cmbSerialNumber2.Items.Add(dr.GetValue(0))
                    End If

                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    Me.cmbSerialNumber1.Items.Add(drLocal.GetValue(0))
                    Me.cmbSerialNumber2.Items.Add(drLocal.GetValue(0))
                End While
                atsLocal.Close()
                SerialListLoad = True
            End If
        Catch
            Return 0
        End Try
        Return 1
    End Function


    Private Function LoadTraceParameters() As Integer
        Try
            Dim cnt As Int32 = 0
            Dim SQLStr As String

            SQLStr = "Select * from Trace"
            If cmbPARM1.Text = "TraceID" Then cmbPARM1.Text = "ID"
            If cmbPARM1.Text = "TraceID" Or cmbPARM1.Text = "TestID" Then
                If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = " & txtPARM1.Text
            Else
                If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
            End If
            If cmbPARM2.Text = "TraceID" Then cmbPARM2.Text = "ID"
            If cmbPARM2.Text = "SerialNumber" Then
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & cmbSerialNumber1.Text & "'"
            ElseIf cmbPARM2.Text = "TraceID" Or cmbPARM2.Text = "TestID" Then
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = " & txtPARM2.Text
            Else
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
            End If
            If cmbPARM3.Text = "TraceID" Then cmbPARM3.Text = "ID"
            If cmbPARM3.Text = "SerialNumber" Then
                If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & cmbSerialNumber2.Text & "'"
            ElseIf cmbPARM3.Text = "TraceID" Or cmbPARM3.Text = "TestID" Then
                If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = " & txtPARM3.Text
            Else
                If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"
            End If

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If Not IsDBNull(dr.GetValue(0)) Then TraceTestID = dr.GetValue(0)
                    If Not IsDBNull(dr.GetValue(3)) Then TraceJob = dr.GetValue(3)
                    If Not IsDBNull(dr.GetValue(4)) Then TracePart = dr.GetValue(4)
                    If Not IsDBNull(dr.GetValue(5)) Then TraceTitle = dr.GetValue(5)
                    If Not IsDBNull(dr.GetValue(6)) Then TraceSN = dr.GetValue(6)
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If Not IsDBNull(drLocal.GetValue(0)) Then TraceTestID = drLocal.GetValue(0)
                    If Not IsDBNull(drLocal.GetValue(3)) Then TraceJob = drLocal.GetValue(3)
                    If Not IsDBNull(drLocal.GetValue(4)) Then TracePart = drLocal.GetValue(4)
                    If Not IsDBNull(drLocal.GetValue(5)) Then TraceTitle = drLocal.GetValue(5)
                    If Not IsDBNull(drLocal.GetValue(6)) Then TraceSN = drLocal.GetValue(6)
                End While
                atsLocal.Close()
                SerialListLoad = True
            End If
        Catch
            Return 0
        End Try
        Return 1
    End Function

    Private Sub cmbPARM1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPARM1.SelectedIndexChanged

        If Not cmbPARM1.GetItemText(cmbPARM1.SelectedItem) = "" Then
            Me.E1.Visible = True
            Me.A1.Visible = True
            Me.txtPARM1.Visible = True
            Me.cmbPARM1.Visible = True
            Me.cmbPARM2.Visible = True
        Else
            Me.E1.Visible = False
            Me.E2.Visible = False
            Me.txtPARM1.Visible = False
            Me.txtPARM2.Visible = False
            Me.txtPARM3.Visible = False
            Me.cmbPARM2.Visible = False
            Me.cmbPARM3.Visible = False
            Me.A1.Visible = False
            Me.E3.Visible = False
        End If
    End Sub
    Private Sub cmbPARM2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPARM2.SelectedIndexChanged
        If Not cmbPARM2.GetItemText(cmbPARM2.SelectedItem) = "" Then
            Me.E1.Visible = True
            Me.E2.Visible = True
            Me.A1.Visible = True
            Me.txtPARM1.Visible = True
            Me.cmbPARM1.Visible = True
            Me.A2.Visible = True
            Me.txtPARM2.Visible = True
            If cmbPARM2.GetItemText(cmbPARM2.SelectedItem).Contains("SerialNumber") Then
                If Not SerialListLoad Then LoadSNLists()
                cmbSerialNumber1.Visible = True
                txtPARM2.Visible = False
            Else
                cmbSerialNumber1.Visible = False
                txtPARM2.Visible = True
            End If
            Me.cmbPARM3.Visible = True
        Else
            cmbSerialNumber1.Visible = False
            Me.E2.Visible = False
            Me.txtPARM2.Visible = False
            Me.txtPARM3.Visible = False
            Me.A2.Visible = False
            Me.cmbPARM3.Visible = False
            Me.E3.Visible = False
        End If



    End Sub
    Private Sub cmbPARM3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbPARM3.SelectedIndexChanged
        If Not cmbPARM3.GetItemText(cmbPARM3.SelectedItem) = "" Then
            Me.E1.Visible = True
            Me.A1.Visible = True
            Me.txtPARM1.Visible = True
            Me.cmbPARM1.Visible = True
            Me.E2.Visible = True
            Me.E3.Visible = True
            Me.txtPARM2.Visible = True
            Me.cmbPARM2.Visible = True
            Me.txtPARM3.Visible = True
            If cmbPARM2.GetItemText(cmbPARM3.SelectedItem).Contains("SerialNumber") Then
                If Not SerialListLoad Then LoadSNLists()
                cmbSerialNumber2.Visible = True
                txtPARM2.Visible = False
            Else
                cmbSerialNumber2.Visible = False
                txtPARM2.Visible = True
            End If
            Me.cmbPARM3.Visible = True
        Else
            cmbSerialNumber2.Visible = False
            Me.E2.Visible = False
            Me.txtPARM3.Visible = False
        End If
    End Sub

    Private Sub cmbExecute_Click(sender As Object, e As EventArgs) Handles cmbExecute.Click
        Dim SQLStr As String
        TraceGrid.Visible = True
        'Example
        'SQLStr = "Select * from Trace where ModelNumber = '" & ModelNumber & "' and JobNumber  = '" & JobNumber & "'"

        SQLStr = "Select * from Trace"
        If cmbPARM1.Text = "TraceID" Then cmbPARM1.Text = "ID"
        If cmbPARM1.Text = "TraceID" Or cmbPARM1.Text = "TestID" Then
            If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = " & txtPARM1.Text
        Else
            If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
        End If
        If cmbPARM2.Text = "TraceID" Then cmbPARM2.Text = "ID"
        If cmbPARM2.Text = "SerialNumber" Then
            If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & cmbSerialNumber1.Text & "'"
        ElseIf cmbPARM2.Text = "TraceID" Or cmbPARM2.Text = "TestID" Then
            If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = " & txtPARM2.Text
        Else
            If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
        End If
        If cmbPARM3.Text = "TraceID" Then cmbPARM3.Text = "ID"
        If cmbPARM3.Text = "SerialNumber" Then
            If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & cmbSerialNumber2.Text & "'"
        ElseIf cmbPARM3.Text = "TraceID" Or cmbPARM3.Text = "TestID" Then
            If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = " & txtPARM3.Text
        Else
            If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"
        End If

        LoadTraceView(SQLStr)
        LoadTestLists()
        LoadSNLists()

    End Sub

    Private Sub Trace1Select_CheckedChanged(sender As Object, e As EventArgs) Handles Trace1Select.CheckedChanged
        If Trace1Select.Checked Then
            Trace1Sel.Visible = True

            Me.lblTestID1.Visible = True
            Me.lblSeialNumber1.Visible = True
            Me.lblWorkStation1.Visible = True
            Me.lblPoints1.Visible = True
            Me.lblDate1.Visible = True

            Me.txtTestID1.Visible = True
            Me.txtSerialNumber1.Visible = True
            Me.txtWorkStation1.Visible = True
            Me.txtPoints1.Visible = True
            Me.txtDate1.Visible = True

            Me.txtTestID1.Text = Nothing
            Me.txtSerialNumber1.Text = Nothing
            Me.txtWorkStation1.Text = Nothing
            Me.txtPoints1.Text = Nothing
            Me.txtDate1.Text = Nothing
        Else
            Trace1Sel.Visible = False

            Me.lblTestID1.Visible = False
            Me.lblSeialNumber1.Visible = False
            Me.lblWorkStation1.Visible = False
            Me.lblPoints1.Visible = False
            Me.lblDate1.Visible = False

            Me.txtTestID1.Visible = False
            Me.txtSerialNumber1.Visible = False
            Me.txtWorkStation1.Visible = False
            Me.txtPoints1.Visible = False
            Me.txtDate1.Visible = False

            Me.txtTestID1.Text = Nothing
            Me.txtSerialNumber1.Text = Nothing
            Me.txtWorkStation1.Text = Nothing
            Me.txtPoints1.Text = Nothing
        End If
    End Sub
    Private Sub Trace1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Trace1Sel.SelectedIndexChanged
        Dim cnt As Int32 = 0
        Dim SQLStr As String
        Try
            SerialListLoad = False
            TraceGrid.Visible = False
            ZG1.Visible = True
            SQLStr = "Select * from Trace"
            If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
            If Not cmbPARM2.Text = Nothing Then
                If cmbPARM2.Text = "SerialNumber" Then
                    SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & cmbSerialNumber1.Text & "'"
                ElseIf cmbPARM2.Text = "TestID" Or cmbPARM2.Text = "SpecID" Then
                    SQLStr = SQLStr & " and " & cmbPARM2.Text & " = " & txtPARM2.Text
                Else
                    SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
                End If
            End If
            If Not cmbPARM3.Text = Nothing Then
                If cmbPARM3.Text = "SerialNumber" Then
                    SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & cmbSerialNumber2.Text & "'"
                ElseIf cmbPARM2.Text = "TestID" Or cmbPARM3.Text = "SpecID" Then
                    SQLStr = SQLStr & " and " & cmbPARM3.Text & " = " & txtPARM3.Text
                Else
                    SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"
                End If
            End If
            If SQLStr.Contains("where") Then
                SQLStr = SQLStr & " and Title = '" & Me.Trace1Sel.GetItemText(Trace1Sel.SelectedItem) & "'"
            Else
                SQLStr = SQLStr & " where Title = '" & Me.Trace1Sel.GetItemText(Trace1Sel.SelectedItem) & "'"
            End If

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    TraceID = (dr.GetValue(0))
                    If Not IsDBNull(dr.GetValue(1)) Then Me.txtTestID1.Text = (dr.GetValue(1))
                    If Not IsDBNull(dr.GetValue(6)) Then Me.txtSerialNumber1.Text = CType(dr.Item(6), String)
                    If Not IsDBNull(dr.GetValue(7)) Then Me.txtWorkStation1.Text = CType(dr.Item(7), String)
                    If Not IsDBNull(dr.GetValue(8)) Then Me.txtPoints1.Text = CType(dr.Item(8), String)
                    If Not IsDBNull(dr.GetValue(9)) Then Me.txtDate1.Text = CType(dr.Item(9), String)
                End While
                ats.Close()

            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    TestID = (drLocal.GetValue(1))
                    If Not IsDBNull(drLocal.GetValue(1)) Then Me.txtTestID1.Text = (drLocal.GetValue(1))
                    If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtSerialNumber1.Text = CType(drLocal.Item(6), String)
                    If Not IsDBNull(drLocal.GetValue(7)) Then Me.txtWorkStation1.Text = CType(drLocal.Item(7), String)
                    If Not IsDBNull(drLocal.GetValue(8)) Then Me.txtPoints1.Text = CType(drLocal.Item(8), String)
                    If Not IsDBNull(drLocal.GetValue(9)) Then Me.txtDate1.Text = CType(drLocal.Item(9), String)
                End While
                atsLocal.Close()

            End If
            TraceID = SQL.GetTraceID(SQLStr, "NetworkTraceData")
            If TraceID > 171666 Then
                GetTracePoints2(TraceID)
            Else
                GetTracePoints(TraceID)
            End If
            Dim i As Integer, x As Double, y As Double
            list1.Clear()
            i = 0
            While Not XArray(i) = Nothing
                x = XArray(i)
                y = YArray(i)
                list1.Add(x, y)
                i = i + 1
            End While



            'Configure graph object default values
            myPane = ZG1.GraphPane
            If IsDBNull(Trace1) Then Trace1.Clear()
            Trace1 = myPane.AddCurve("T1:" & Me.Trace1Sel.GetItemText(Trace1Sel.SelectedItem), list1, Color.Chartreuse, SymbolType.None)
            ' Fill the symbols with white
            Trace1.Symbol.Fill = New Fill(Color.White)


            myPane.Title.Text = Me.Trace1Sel.GetItemText(Trace1Sel.SelectedItem)
            myPane.XAxis.Title.Text = "Frequency (MHz)"
            If list1.Count > 0 Then ScaleGraph()

        Catch

        End Try
    End Sub

    Private Sub Trace2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Trace2Sel.SelectedIndexChanged
        Dim cnt As Int32 = 0
        Dim SQLStr As String
        Try
            SerialListLoad = False
            TraceGrid.Visible = False
            ZG1.Visible = True
            SQLStr = "Select * from Trace"
            If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
            If Not cmbPARM2.Text = Nothing Then
                If cmbPARM2.Text = "SerialNumber" Then
                    SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & cmbSerialNumber1.Text & "'"
                Else
                    SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
                End If
            End If
            If Not cmbPARM3.Text = Nothing Then
                If cmbPARM3.Text = "SerialNumber" Then
                    SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & cmbSerialNumber2.Text & "'"
                Else
                    SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"
                End If
            End If
            If SQLStr.Contains("where") Then
                SQLStr = SQLStr & " and Title = '" & Me.Trace2Sel.GetItemText(Trace2Sel.SelectedItem) & "'"
            Else
                SQLStr = SQLStr & " where Title = '" & Me.Trace2Sel.GetItemText(Trace2Sel.SelectedItem) & "'"
            End If

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    TraceID = (dr.GetValue(0))
                    If Not IsDBNull(dr.GetValue(1)) Then Me.txtTestID2.Text = (dr.GetValue(1))
                    If Not IsDBNull(dr.GetValue(6)) Then Me.txtSerialNumber2.Text = CType(dr.Item(6), String)
                    If Not IsDBNull(dr.GetValue(7)) Then Me.txtWorkStation2.Text = CType(dr.Item(7), String)
                    If Not IsDBNull(dr.GetValue(8)) Then Me.txtPoints2.Text = CType(dr.Item(8), String)
                    If Not IsDBNull(dr.GetValue(9)) Then Me.txtDate2.Text = CType(dr.Item(9), String)
                End While
                ats.Close()

            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    TraceID = (drLocal.GetValue(0))
                    If Not IsDBNull(drLocal.GetValue(1)) Then Me.txtTestID2.Text = (drLocal.GetValue(1))
                    If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtSerialNumber2.Text = CType(drLocal.Item(6), String)
                    If Not IsDBNull(drLocal.GetValue(7)) Then Me.txtWorkStation2.Text = CType(drLocal.Item(7), String)
                    If Not IsDBNull(drLocal.GetValue(8)) Then Me.txtPoints2.Text = CType(drLocal.Item(8), String)
                    If Not IsDBNull(drLocal.GetValue(9)) Then Me.txtDate2.Text = CType(drLocal.Item(9), Date)
                End While
                atsLocal.Close()

            End If
            TraceID = SQL.GetTraceID(SQLStr, "NetworkTraceData")
            If TraceID > 171666 Then
                GetTracePoints2(TraceID)
            Else
                GetTracePoints(TraceID)
            End If
            Dim i As Integer, x As Double, y As Double
            list2.Clear()
            i = 0
            While Not XArray(i) = Nothing
                x = XArray(i)
                y = YArray(i)
                list2.Add(x, y)
                i = i + 1
            End While

            'Configure graph object default values
            If IsDBNull(Trace2) Then Trace2.Clear()
            myPane = ZG1.GraphPane
            ' Generate a curve for Tanalyzer
            'Trace1.Clear()
            Trace2 = myPane.AddCurve("T2:" & Me.Trace2Sel.GetItemText(Trace2Sel.SelectedItem), list2, Color.CornflowerBlue, SymbolType.None)
            ' Fill the symbols with white
            Trace2.Symbol.Fill = New Fill(Color.White)

            myPane.Title.Text = Me.Trace2Sel.GetItemText(Trace2Sel.SelectedItem)
            myPane.XAxis.Title.Text = "Frequency (MHz)"
            ScaleGraph()

        Catch

        End Try
    End Sub
    Private Sub Trace3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Trace3Sel.SelectedIndexChanged
        Dim cnt As Int32 = 0
        Dim SQLStr As String
        Try
            SerialListLoad = False
            TraceGrid.Visible = False
            ZG1.Visible = True
            SQLStr = "Select * from Trace"
            If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
            If Not cmbPARM2.Text = Nothing Then
                If cmbPARM2.Text = "SerialNumber" Then
                    SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & cmbSerialNumber1.Text & "'"
                Else
                    SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
                End If
            End If
            If Not cmbPARM3.Text = Nothing Then
                If cmbPARM3.Text = "SerialNumber" Then
                    SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & cmbSerialNumber2.Text & "'"
                Else
                    SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"
                End If
            End If
            If SQLStr.Contains("where") Then
                SQLStr = SQLStr & " and Title = '" & Me.Trace3Sel.GetItemText(Trace3Sel.SelectedItem) & "'"
            Else
                SQLStr = SQLStr & " where Title = '" & Me.Trace3Sel.GetItemText(Trace3Sel.SelectedItem) & "'"
            End If

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    TraceID = (dr.GetValue(0))
                    If Not IsDBNull(dr.GetValue(1)) Then Me.txtTestID3.Text = (dr.GetValue(1))
                    If Not IsDBNull(dr.GetValue(6)) Then Me.txtSerialNumber3.Text = CType(dr.Item(6), String)
                    If Not IsDBNull(dr.GetValue(7)) Then Me.txtWorkStation3.Text = CType(dr.Item(7), String)
                    If Not IsDBNull(dr.GetValue(8)) Then Me.txtPoints3.Text = CType(dr.Item(8), String)
                    If Not IsDBNull(dr.GetValue(9)) Then Me.txtDate3.Text = CType(dr.Item(9), String)
                End While
                ats.Close()

            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    TraceID = (drLocal.GetValue(0))
                    If Not IsDBNull(drLocal.GetValue(1)) Then Me.txtTestID3.Text = (drLocal.GetValue(1))
                    If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtSerialNumber3.Text = CType(drLocal.Item(6), String)
                    If Not IsDBNull(drLocal.GetValue(7)) Then Me.txtWorkStation3.Text = CType(drLocal.Item(6), String)
                    If Not IsDBNull(drLocal.GetValue(8)) Then Me.txtPoints3.Text = CType(drLocal.Item(7), String)
                    If Not IsDBNull(drLocal.GetValue(9)) Then Me.txtDate3.Text = CType(drLocal.Item(8), String)
                End While
                atsLocal.Close()

            End If
            TraceID = SQL.GetTraceID(SQLStr, "NetworkTraceData")
            If TraceID > 171666 Then
                GetTracePoints2(TraceID)
            Else
                GetTracePoints(TraceID)
            End If
            Dim i As Integer, x As Double, y As Double
            list3.Clear()
            i = 0
            While Not XArray(i) = Nothing
                x = XArray(i)
                y = YArray(i)
                list3.Add(x, y)
                i = i + 1
            End While

            'Configure graph object default values
            If IsDBNull(Trace3) Then Trace3.Clear()
            myPane = ZG1.GraphPane
            ' Generate a curve for Tanalyzer
            'Trace1.Clear()
            Trace3 = myPane.AddCurve("T3:" & Me.Trace3Sel.GetItemText(Trace3Sel.SelectedItem), list3, Color.LightGreen, SymbolType.None)
            ' Fill the symbols with white
            Trace1.Symbol.Fill = New Fill(Color.White)

            myPane.Title.Text = Me.Trace3Sel.GetItemText(Trace3Sel.SelectedItem)
            myPane.XAxis.Title.Text = "Frequency (MHz)"
            ScaleGraph()

        Catch

        End Try
    End Sub
    Private Sub Trace4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Trace4Sel.SelectedIndexChanged
        Dim cnt As Int32 = 0
        Dim SQLStr As String
        Try
            SerialListLoad = False
            TraceGrid.Visible = False
            ZG1.Visible = True
            SQLStr = "Select * from Trace"
            If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
            If Not cmbPARM2.Text = Nothing Then
                If cmbPARM2.Text = "SerialNumber" Then
                    SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & cmbSerialNumber1.Text & "'"
                Else
                    SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
                End If
            End If
            If Not cmbPARM3.Text = Nothing Then
                If cmbPARM3.Text = "SerialNumber" Then
                    SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & cmbSerialNumber2.Text & "'"
                Else
                    SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"
                End If
            End If
            If SQLStr.Contains("where") Then
                SQLStr = SQLStr & " and Title = '" & Me.Trace4Sel.GetItemText(Trace4Sel.SelectedItem) & "'"
            Else
                SQLStr = SQLStr & " where Title = '" & Me.Trace4Sel.GetItemText(Trace4Sel.SelectedItem) & "'"
            End If

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    TraceID = (dr.GetValue(0))
                    If Not IsDBNull(dr.GetValue(1)) Then Me.txtTestID4.Text = (dr.GetValue(1))
                    If Not IsDBNull(dr.GetValue(6)) Then Me.txtSerialNumber4.Text = CType(dr.Item(6), String)
                    If Not IsDBNull(dr.GetValue(7)) Then Me.txtWorkStation4.Text = CType(dr.Item(7), String)
                    If Not IsDBNull(dr.GetValue(8)) Then Me.txtPoints4.Text = CType(dr.Item(8), String)
                    If Not IsDBNull(dr.GetValue(9)) Then Me.txtDate4.Text = CType(dr.Item(9), String)
                End While
                ats.Close()

            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & AccessDatabaseFolder("NetworkTraceData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    TraceID = (drLocal.GetValue(0))
                    If Not IsDBNull(drLocal.GetValue(1)) Then Me.txtTestID4.Text = (drLocal.GetValue(1))
                    If Not IsDBNull(drLocal.GetValue(6)) Then Me.txtSerialNumber4.Text = CType(drLocal.Item(6), String)
                    If Not IsDBNull(drLocal.GetValue(7)) Then Me.txtWorkStation4.Text = CType(drLocal.Item(7), String)
                    If Not IsDBNull(drLocal.GetValue(8)) Then Me.txtPoints4.Text = CType(drLocal.Item(8), String)
                    If Not IsDBNull(drLocal.GetValue(9)) Then Me.txtDate1.Text = CType(drLocal.Item(9), String)
                End While
                atsLocal.Close()

            End If
            TraceID = SQL.GetTraceID(SQLStr, "NetworkTraceData")
            If TraceID > 171666 Then
                GetTracePoints2(TraceID)
            Else
                GetTracePoints(TraceID)
            End If
            Dim i As Integer, x As Double, y As Double
            list4.Clear()
            i = 0
            While Not XArray(i) = Nothing
                x = XArray(i)
                y = YArray(i)
                list4.Add(x, y)
                i = i + 1
            End While

            'Configure graph object default values
            myPane = ZG1.GraphPane
            ' Generate a curve for Tanalyzer
            'Trace1.Clear()
            Trace4 = myPane.AddCurve("T4:" & Me.Trace4Sel.GetItemText(Trace4Sel.SelectedItem), list4, Color.LightCoral, SymbolType.None)
            ' Fill the symbols with white
            Trace1.Symbol.Fill = New Fill(Color.White)

            myPane.Title.Text = Me.Trace4Sel.GetItemText(Trace4Sel.SelectedItem)
            myPane.XAxis.Title.Text = "Frequency (MHz)"
            ScaleGraph()

        Catch

        End Try
    End Sub

    Private Sub Trace2Select_CheckedChanged(sender As Object, e As EventArgs) Handles Trace2Select.CheckedChanged
        If Trace2Select.Checked Then
            Trace2Sel.Visible = True

            Me.lblTestD2.Visible = True
            Me.lblSeialNumber2.Visible = True
            Me.lblWorkStation2.Visible = True
            Me.lblPoints2.Visible = True
            Me.lblDate2.Visible = True

            Me.txtTestID2.Visible = True
            Me.txtSerialNumber2.Visible = True
            Me.txtWorkStation2.Visible = True
            Me.txtPoints2.Visible = True
            Me.txtDate2.Visible = True

            Me.txtTestID2.Text = Nothing
            Me.txtSerialNumber2.Text = Nothing
            Me.txtWorkStation2.Text = Nothing
            Me.txtPoints2.Text = Nothing
            Me.txtDate2.Text = Nothing
        Else
            Trace2Sel.Visible = False

            Me.lblTestD2.Visible = False
            Me.lblSeialNumber2.Visible = False
            Me.lblWorkStation2.Visible = False
            Me.lblPoints2.Visible = False
            Me.lblDate2.Visible = False

            Me.txtTestID2.Visible = False
            Me.txtSerialNumber2.Visible = False
            Me.txtWorkStation2.Visible = False
            Me.txtPoints2.Visible = False
            Me.txtDate2.Visible = False

            Me.txtTestID2.Text = Nothing
            Me.txtSerialNumber2.Text = Nothing
            Me.txtWorkStation2.Text = Nothing
            Me.txtPoints2.Text = Nothing
        End If
    End Sub

    Private Sub Trace3Select_CheckedChanged(sender As Object, e As EventArgs) Handles Trace3Select.CheckedChanged

        If Trace3Select.Checked Then
            Trace3Sel.Visible = True

            Me.lblTestID3.Visible = True
            Me.lblSeialNumber3.Visible = True
            Me.lblWorkStation3.Visible = True
            Me.lblPoints3.Visible = True
            Me.lblDate3.Visible = True

            Me.txtTestID3.Visible = True
            Me.txtSerialNumber3.Visible = True
            Me.txtWorkStation3.Visible = True
            Me.txtPoints3.Visible = True
            Me.txtDate3.Visible = True

            Me.txtTestID3.Text = Nothing
            Me.txtSerialNumber3.Text = Nothing
            Me.txtWorkStation3.Text = Nothing
            Me.txtPoints3.Text = Nothing
            Me.txtDate3.Text = Nothing
        Else
            Trace3Sel.Visible = False

            Me.lblTestID3.Visible = False
            Me.lblSeialNumber3.Visible = False
            Me.lblWorkStation3.Visible = False
            Me.lblPoints3.Visible = False
            Me.lblDate3.Visible = False

            Me.txtTestID3.Visible = False
            Me.txtSerialNumber3.Visible = False
            Me.txtWorkStation3.Visible = False
            Me.txtPoints3.Visible = False
            Me.txtDate3.Visible = False

            Me.txtTestID3.Text = Nothing
            Me.txtSerialNumber3.Text = Nothing
            Me.txtWorkStation3.Text = Nothing
            Me.txtPoints3.Text = Nothing
        End If
    End Sub

    Private Sub Trace4Select_CheckedChanged(sender As Object, e As EventArgs) Handles Trace4Select.CheckedChanged

        If Trace4Select.Checked Then
            Trace4Sel.Visible = True

            Me.lblTestID4.Visible = True
            Me.lblSeialNumber4.Visible = True
            Me.lblWorkStation4.Visible = True
            Me.lblPoints4.Visible = True
            Me.lblDate4.Visible = True

            Me.txtTestID4.Visible = True
            Me.txtSerialNumber4.Visible = True
            Me.txtWorkStation4.Visible = True
            Me.txtPoints4.Visible = True
            Me.txtDate4.Visible = True

            Me.txtTestID4.Text = Nothing
            Me.txtSerialNumber4.Text = Nothing
            Me.txtWorkStation4.Text = Nothing
            Me.txtPoints4.Text = Nothing

            Me.txtDate4.Text = Nothing
        Else
            Trace4Sel.Visible = False

            Me.lblTestID4.Visible = False
            Me.lblSeialNumber4.Visible = False
            Me.lblWorkStation4.Visible = False
            Me.lblPoints4.Visible = False
            Me.lblDate4.Visible = False

            Me.txtTestID4.Visible = False
            Me.txtSerialNumber4.Visible = False
            Me.txtWorkStation4.Visible = False
            Me.txtPoints4.Visible = False
            Me.txtDate4.Visible = False

            Me.txtTestID4.Text = Nothing
            Me.txtSerialNumber4.Text = Nothing
            Me.txtWorkStation4.Text = Nothing
            Me.txtPoints4.Text = Nothing
        End If
    End Sub


    Private Sub ScaleGraph()
        'Dim x As Integer
        Try
            myPane.YAxis.Scale.Min = YArray.Min + 1
            myPane.YAxis.Scale.Max = YArray.Max - 1
            myPane.XAxis.Scale.Min = XArray(0)
            If Not list1.Count = 0 Then myPane.XAxis.Scale.Max = XArray(list1.Count - 1)
            ' myPane.YAxis.Scale.BaseTic = (cxrGraph.YAxis.MaxValue - cxrGraph.YAxis.MinValue) / 10
            'cxrGraph.XAxis.MajorTickSpacing = (cxrGraph.XAxis.MaxValue - cxrGraph.XAxis.MinValue) / 10

            'cxrGraph.YAxis.XIntercept = cxrGraph.YAxis.MinValue * 2
            'Trace1.Clear()
            'trace2.Clear()

            'For x = UBound(FreqAry) To 0
            '    'trace1.RemovePoint(x)
            '    trace2.RemovePoint(x)
            'Next
            'For x = 0 To UBound(FreqAry)
            '    'trace1.AddPoint(NMLxAry(x), NMLyAry(x))
            '    trace2.AddPoint(FreqAry(x), dyAry(x))
            'Next x
            '' Tell ZedGraph to calculate the axis ranges
            ' Note that you MUST call this after enabling IsAutoScrollRange, since AxisChange() sets
            ' up the proper scrolling parameters
            ZG1.AxisChange()


            ZG1.Invalidate()       ' Tell ZedGraph to calculate the axis ranges
            ' Note that you MUST call this after enabling IsAutoScrollRange, since AxisChange() sets
            ' up the proper scrolling parameters
            ZG1.AxisChange()


            ZG1.Invalidate()
        Catch
        End Try
    End Sub


    Public Sub PrintTraces()
        Dim Row As Long
        Dim TopFolder As String
        Dim SubFolder As String
        Dim TraceID1 As Long
        Dim TraceID2 As Long
        Dim TraceID3 As Long
        Dim TraceID4 As Long
        Dim XData1(Points) As Double
        Dim YData1(Points) As Double
        Dim XData2(Points) As Double
        Dim YData2(Points) As Double
        Dim XData3(Points) As Double
        Dim YData3(Points) As Double
        Dim XData4(Points) As Double
        Dim YData4(Points) As Double

        Try
            SpecType = SQL.GetSpecification("SpecType")
            If SpecType = "N/A" Then Exit Sub


            ExcelReports.StartupReport(ExcelTemplatePath, "TraceData.xls")
            ExcelReports.DataToCells("C5", TraceSN)
            ExcelReports.DataToCells("F2", TraceJob)
            ExcelReports.DataToCells("F3", TracePart)
            ExcelReports.DataToCells("F4", TraceTitle)

            Row = 10
            If Trace1Select.Checked Then
                ExcelReports.DataToCells("A8", Trace1Sel.Text)
                TraceID1 = GetTraceIDByTitle(Trace1Sel.Text, SerialNumber, TraceJob)
                GetTracePoints(TraceID1)
                XData1 = XArray
                YData1 = YArray
            End If
            If Trace2Select.Checked Then
                ExcelReports.DataToCells("C8", Trace2Sel.Text)
                TraceID2 = GetTraceIDByTitle(Trace2Sel.Text, SerialNumber, TraceJob)
                GetTracePoints(TraceID2)
                XData2 = XArray
                YData2 = YArray
            End If
            If Trace3Select.Checked Then
                ExcelReports.DataToCells("E8", Trace3Sel.Text)
                TraceID3 = GetTraceIDByTitle(Trace3Sel.Text, SerialNumber, TraceJob)
                GetTracePoints(TraceID3)
                XData3 = XArray
                YData3 = YArray
            End If
            If Trace4Select.Checked Then
                ExcelReports.DataToCells("G8", Trace4Sel.Text)
                TraceID4 = GetTraceIDByTitle(Trace4Sel.Text, SerialNumber, TraceJob)
                GetTracePoints(TraceID4)
                XData4 = XArray
                YData4 = YArray
            End If

            If TraceID1 = -1 Then
                MYMsgBox("No trace available")
                Exit Sub
            End If


            For n = 0 To Points - 1
                If TraceID1 > 0 Then
                    ExcelReports.DataToCells("A" & Row, XData1(n))
                    ExcelReports.DataToCells("B" & Row, YData1(n))
                End If
                If TraceID2 > 0 Then
                    ExcelReports.DataToCells("C" & Row, XData2(n))
                    ExcelReports.DataToCells("D" & Row, YData2(n))
                End If
                If TraceID3 > 0 Then
                    ExcelReports.DataToCells("E" & Row, XData3(n))
                    ExcelReports.DataToCells("F" & Row, YData3(n))
                End If
                If TraceID4 > 0 Then
                    ExcelReports.DataToCells("G" & Row, XData4(n))
                    ExcelReports.DataToCells("H" & Row, YData4(n))
                End If
                Row = Row + 1
            Next n

            TopFolder = ""
            If SpecType.Contains("90 DEGREE COUPLER") Then TopFolder = "90_Degree\"
            If SpecType.Contains("90 DEGREE COUPLER SMD") Then TopFolder = "90_Degree_SMD\"
            If SpecType.Contains("DIRECTIONAL COUPLER") Then TopFolder = "Directional_Couplers\"
            If SpecType.Contains("DIRECTIONAL COUPLER SMD") Then TopFolder = "Directional_Couplers_SMD\"
            If SpecType.Contains("COMBINER/DIVIDER") Then TopFolder = "Combiner-Divider\"
            If SpecType.Contains("COMBINER/DIVIDER SMD") Then TopFolder = "Combiner-Divider_SMD\"

            If NetworkAccess Then
                'Network Save
                If Not System.IO.Directory.Exists(TestDataPath & TopFolder) Then System.IO.Directory.CreateDirectory(TestDataPath & TopFolder)
                SubFolder = TestDataPath & TopFolder & TracePart & "-" & TraceJob & "\"
            Else
                'Local Save
                If Not System.IO.Directory.Exists("C:\ATE Data\" & TopFolder) Then System.IO.Directory.CreateDirectory("C:\ATE Data\" & TopFolder)
                SubFolder = "C:\ATE Data\" & TopFolder & TracePart & "-" & TraceJob & "\"

            End If
            If Not System.IO.Directory.Exists(SubFolder) = "" Then System.IO.Directory.CreateDirectory((SubFolder))
            ExcelReports.SaveAs(SubFolder & "TestData " & "TraceData.xls")

        Catch
            MYMsgBox("An Error has Occured In The Trace Data Report" & vbCr & "Report This Error To AutomatedTestSolutions@Gmail.com" & vbCr & "Error Details :-" & vbCr & "Error Number : " & Err.Number & vbCr & "Error Description : " & Err.Description, vbCritical, "FlexGrid Example")
        End Try

    End Sub
  

  
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PrintTraces()
    End Sub
End Class