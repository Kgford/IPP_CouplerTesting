Imports System.Data
Imports System.Data.SqlClient

Public Class EffeciencyView

    Public Sub New()
        MyBase.New()
        Dim SQLStr As String
        Try
            InitializeComponent()

            SQLStr = "Select * from Effeciency where RunStatus <> 'Archived'"
            LoadEFFData(SQLStr)
            EffeciencyChart()

            cmbSelect.Text = "Select"

        Catch
        End Try
    End Sub
    Private Sub LoadEFFData(SQLStr As String)
        Dim NameStr(9) As String

        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        NameStr(0) = "ID"
        NameStr(1) = "WorkStation"
        NameStr(2) = "JobNumber"
        NameStr(3) = "PartNumber"
        NameStr(4) = "Operator"
        NameStr(5) = "ActiveDate"
        NameStr(6) = "TotalUUTs"
        NameStr(7) = "CompleteUUTs"
        NameStr(8) = "EffeciencyStatus"
        NameStr(9) = "RunStatus"
       

        Try

            For x = 0 To 9
                Dim col As New DataGridViewTextBoxColumn

                col.DataPropertyName = NameStr(x)
                col.HeaderText = NameStr(x)
                col.Name = NameStr(x)
                DataGridView1.Columns.Add(col)
            Next

            'SEARCH BY MODEL AND JOB
            'SQLStr = "Select * from ModelConfig where ModelNumber = '" & ModelNumber & "' and JobNumber  = '" & JobNumber & "'"

            If CheckforRow(SQLStr, "Effeciency") > 0 Then GoTo GotIt

GotIt:
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    Dim Count As Integer = DataGridView1.Rows.Add()
                    If Not IsDBNull(dr.Item(0)) Then DataGridView1.Rows.Item(Count).Cells(0).Value = CType(dr.Item(0), String)
                    If Not IsDBNull(dr.Item(1)) Then DataGridView1.Rows.Item(Count).Cells(1).Value = CType(dr.Item(1), String)
                    If Not IsDBNull(dr.Item(2)) Then DataGridView1.Rows.Item(Count).Cells(2).Value = CType(dr.Item(2), String)
                    If Not IsDBNull(dr.Item(3)) Then DataGridView1.Rows.Item(Count).Cells(3).Value = CType(dr.Item(3), String)
                    If Not IsDBNull(dr.Item(4)) Then DataGridView1.Rows.Item(Count).Cells(4).Value = CType(dr.Item(4), String)
                    If Not IsDBNull(dr.Item(5)) Then DataGridView1.Rows.Item(Count).Cells(5).Value = CType(dr.Item(5), String)
                    If Not IsDBNull(dr.Item(6)) Then DataGridView1.Rows.Item(Count).Cells(6).Value = CType(dr.Item(6), String)
                    If Not IsDBNull(dr.Item(7)) Then DataGridView1.Rows.Item(Count).Cells(7).Value = CType(dr.Item(7), String)
                    If Not IsDBNull(dr.Item(8)) Then DataGridView1.Rows.Item(Count).Cells(8).Value = CType(dr.Item(8), String)
                    If Not IsDBNull(dr.Item(9)) Then DataGridView1.Rows.Item(Count).Cells(9).Value = CType(dr.Item(9), String)

                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("Effeciency")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    Dim Count As Integer = DataGridView1.Rows.Add()

                    If Not IsDBNull(drLocal.Item(0)) Then DataGridView1.Rows.Item(Count).Cells(0).Value = CType(drLocal.Item(0), String)
                    If Not IsDBNull(drLocal.Item(1)) Then DataGridView1.Rows.Item(Count).Cells(1).Value = CType(drLocal.Item(1), String)
                    If Not IsDBNull(drLocal.Item(2)) Then DataGridView1.Rows.Item(Count).Cells(2).Value = CType(drLocal.Item(2), String)
                    If Not IsDBNull(drLocal.Item(3)) Then DataGridView1.Rows.Item(Count).Cells(3).Value = CType(drLocal.Item(3), String)
                    If Not IsDBNull(drLocal.Item(4)) Then DataGridView1.Rows.Item(Count).Cells(4).Value = CType(drLocal.Item(4), String)
                    If Not IsDBNull(drLocal.Item(5)) Then DataGridView1.Rows.Item(Count).Cells(5).Value = CType(drLocal.Item(5), String)
                    If Not IsDBNull(drLocal.Item(6)) Then DataGridView1.Rows.Item(Count).Cells(6).Value = CType(drLocal.Item(6), String)
                    If Not IsDBNull(drLocal.Item(7)) Then DataGridView1.Rows.Item(Count).Cells(7).Value = CType(drLocal.Item(7), String)
                    If Not IsDBNull(drLocal.Item(8)) Then DataGridView1.Rows.Item(Count).Cells(8).Value = CType(drLocal.Item(8), String)
                    If Not IsDBNull(drLocal.Item(9)) Then DataGridView1.Rows.Item(Count).Cells(9).Value = CType(drLocal.Item(9), String)
                    

                End While
                atsLocal.Close()
            End If
            '*******************************************************************************
        Catch
        End Try
    End Sub

    Private Sub Tables_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

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
            Me.cmbPARM2.Visible = True
            Me.cmbPARM3.Visible = True
        Else
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
            Me.cmbPARM3.Visible = True
        Else
            Me.E2.Visible = False
            Me.txtPARM3.Visible = False
        End If
    End Sub

    Private Sub cmbExecute_Click(sender As Object, e As EventArgs) Handles cmbExecute.Click
        Dim SQLStr As String

        'Example
        'SQLStr = "Select * TestData where JobNumber = '" & ModelNumber & "' and JobNumber  = '" & JobNumber & "'"
        If cmbSelect.Text = "Delete" Then
            If MessageBox.Show("Are you Sure you want to Delete", "****CANNOT BE REVERSED!!!!!****", MessageBoxButtons.YesNo) = vbYes Then
                GoTo DoIt
            Else
                GoTo DontDoIt
            End If
        End If
DoIt:
        If cmbSelect.Text = "Delete" Then
            SQLStr = cmbSelect.Text & " from Effeciency"
        Else
            SQLStr = cmbSelect.Text & " * from Effeciency"
        End If
        If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
        If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
        If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"
        LoadEFFData(SQLStr)
DontDoIt:


    End Sub


    Private Sub EffeciencyChart()
        Dim SQLStr As String
        Dim Count As Integer = 0
        Dim x As Integer = 0
        Dim y As Integer = 0
        Dim Oper(15) As String
       
        Dim AvgSeries(15) As String
        Dim NowSeries(15) As String
        Dim EffArray(3) As String
        Dim XLOC As Double
        Dim YMax As Double = 0
        Try

            SQLStr = "Select DISTINCT Operator from Effeciency where RunStatus <> 'Archived'"
            If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
            If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
            If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"

            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    If Not IsDBNull(dr.Item(0)) Then Chart1.Series.Add(CStr(dr.Item(0)))
                    If Not IsDBNull(dr.Item(0)) Then AvgSeries(Count) = CStr(dr.Item(0))
                    If Not IsDBNull(dr.Item(0)) Then Oper(Count) = CStr(dr.Item(0))
                    If Not IsDBNull(dr.Item(0)) Then Count = Count + 1
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("Effeciency")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    If Not IsDBNull(drLocal.Item(0)) Then Chart1.Series.Add(CStr(drLocal.Item(0)))
                    If Not IsDBNull(drLocal.Item(0)) Then AvgSeries(Count) = CStr(drLocal.Item(0))
                    If Not IsDBNull(drLocal.Item(0)) Then Oper(Count) = CStr(drLocal.Item(0))
                    If Not IsDBNull(drLocal.Item(0)) Then Count = Count + 1
                End While
                atsLocal.Close()
            End If
            Dim OperAve(Count) As Double
            Dim OperNow(Count) As Double
            If Count > 0 Then

                For x = 0 To Count - 1 Step 1
                    y = 0
                    '********************************* Operator Average***********************************************
                    SQLStr = "Select * from Effeciency where Operator = '" & Oper(x) & "' and RunStatus = 'Complete'"
                    OperAve(x) = 0
                    If SQLAccess Then
                        Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                        Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                        ats.Open()
                        System.Threading.Thread.Sleep(10)
                        Dim dr As SqlDataReader = cmd.ExecuteReader()
                        While Not dr.Read = Nothing
                            If Not IsDBNull(dr.Item(8)) Then EffArray = Split(CStr(dr.Item(8)), "%")
                            If Not IsDBNull(dr.Item(8)) Then OperAve(x) = OperAve(x) + CDbl(EffArray(0))
                            y = y + 1
                        End While
                        ats.Close()
                    Else
                        Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("Effeciency")
                        Dim atsLocal As New OleDb.OleDbConnection
                        atsLocal.ConnectionString = strConnectionString
                        Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                        atsLocal.Open()
                        System.Threading.Thread.Sleep(10)
                        Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                        While Not drLocal.Read = Nothing
                            If Not IsDBNull(drLocal.Item(8)) Then EffArray = Split(CStr(drLocal.Item(8)), "%")
                            If Not IsDBNull(drLocal.Item(8)) Then OperAve(x) = OperAve(x) + CDbl(EffArray(0))
                            y = y + 1
                        End While
                        atsLocal.Close()
                    End If
                    'If x = 0 Then
                    '    XLOC = x
                    'Else
                    '    XLOC = x + 1
                    'End If
                    XLOC = x
                    OperAve(x) = OperAve(x) / y
                    Chart1.Series(AvgSeries(x)).Points.AddXY(XLOC, OperAve(x))
                    Chart1.Series(AvgSeries(x)).AxisLabel = "AVG"
                    If OperAve(x) > YMax Then YMax = OperAve(x)
                    '********************************* Operator Average***********************************************

                    '*********************************** Operator Now *************************************************
                    SQLStr = "Select* from Effeciency where Operator = '" & Oper(x) & "' and RunStatus = 'Running'"
                    OperAve(Count) = 0
                    If SQLAccess Then
                        Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                        Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                        ats.Open()
                        System.Threading.Thread.Sleep(10)
                        Dim dr As SqlDataReader = cmd.ExecuteReader()
                        While Not dr.Read = Nothing
                            If Not IsDBNull(dr.Item(8)) Then EffArray = Split(CStr(dr.Item(8)), "%")
                            If Not IsDBNull(dr.Item(8)) Then OperNow(x) = CDbl(EffArray(0))
                        End While
                        ats.Close()
                    Else
                        Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("Effeciency")
                        Dim atsLocal As New OleDb.OleDbConnection
                        atsLocal.ConnectionString = strConnectionString
                        Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                        atsLocal.Open()
                        System.Threading.Thread.Sleep(10)
                        Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                        While Not drLocal.Read = Nothing
                            If Not IsDBNull(drLocal.Item(8)) Then EffArray = Split(CStr(drLocal.Item(8)), "%")
                            If Not IsDBNull(drLocal.Item(8)) Then OperNow(x) = CDbl(EffArray(0))
                        End While
                        atsLocal.Close()
                    End If
                    '*********************************** Operator Now *************************************************
                    Chart1.Series(AvgSeries(x)).Points.AddXY(XLOC + 0.3, OperNow(x))
                    Chart1.Series(AvgSeries(x)).AxisLabel = "Now"

                    If OperNow(x) > YMax Then YMax = OperNow(x)

                Next x
                ' Chart1.ChartAreas(0).AxisX.Maximum = (x * 2) - 1
                Chart1.ChartAreas(0).AxisY.Maximum = YMax + (YMax / 10)
                Chart1.ChartAreas(0).AxisY.Title = "Effeciency %"
                Chart1.ChartAreas(0).AxisY.TitleFont = New Font("Arial", 20, FontStyle.Bold)
                Chart1.ChartAreas(0).AxisX.TitleFont = New Font("Arial", 10, FontStyle.Bold)

            End If

        Catch
        End Try
    End Sub


   
    Private Sub bTTsEARCH_Click(sender As Object, e As EventArgs) Handles bTTsEARCH.Click
        If bTTsEARCH.Text = "DATABASE" Then
            bTTsEARCH.Text = "CHART"
            cmbExecute.Visible = True
            DataGridView1.Visible = True
            Chart1.Visible = False
        Else
            bTTsEARCH.Text = "DATABASE"
            cmbExecute.Visible = False
            DataGridView1.Visible = False
            Chart1.Visible = True
        End If
    End Sub
End Class