Imports System.Data
Imports System.Data.SqlClient

Public Class SpecDataView

    Public Sub New()
        MyBase.New()
        Dim SQLStr As String
        Try
            InitializeComponent()

            SQLStr = "Select * from Specifications"
            LoadTestData(SQLStr)

            cmbSelect.Text = "Select"

        Catch
        End Try
    End Sub
    Private Sub LoadTestData(SQLStr As String)
        Dim NameStr(32) As String

        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        NameStr(0) = "SpecID"
        NameStr(1) = "SpecType"
        NameStr(2) = "JobNumber"
        NameStr(3) = "PartNumber"
        NameStr(4) = "Title"
        NameStr(5) = "Quantity"
        NameStr(6) = "StartFreqMHz"
        NameStr(7) = "StopFreqMHz"
        NameStr(8) = "CutOffFreqMHz"
        NameStr(9) = "OutputPortNumber"
        NameStr(10) = "VSWR"
        NameStr(11) = "InsertionLoss"
        NameStr(12) = "Isolation"
        NameStr(13) = "Isolation2"
        NameStr(14) = "AmplitudeBalance"
        NameStr(15) = "Coupling"
        NameStr(16) = "CoupPlusMinus"
        NameStr(17) = "Directivity"
        NameStr(18) = "PhaseBalance"
        NameStr(19) = "CoupledFlatness"
        NameStr(20) = "Power"
        NameStr(21) = "Temperature"
        NameStr(22) = "Offset1"
        NameStr(23) = "Offset2"
        NameStr(24) = "Offset3"
        NameStr(25) = "Offset4"
        NameStr(26) = "Offset5"
        NameStr(27) = "Test1"
        NameStr(28) = "Test2"
        NameStr(29) = "Test3"
        NameStr(30) = "Test4"
        NameStr(31) = "Test5"
        NameStr(32) = "PartsPerHour"


        Try

            For x = 0 To 32
                Dim col As New DataGridViewTextBoxColumn

                col.DataPropertyName = NameStr(x)
                col.HeaderText = NameStr(x)
                col.Name = NameStr(x)
                DataGridView1.Columns.Add(col)
            Next

            'SEARCH BY MODEL AND JOB
            'SQLStr = "Select * from ModelConfig where ModelNumber = '" & ModelNumber & "' and JobNumber  = '" & JobNumber & "'"

            If CheckforRow(SQLStr, "NetworkSpecs") > 0 Then GoTo GotIt

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
                    If Not IsDBNull(dr.Item(10)) Then DataGridView1.Rows.Item(Count).Cells(10).Value = CType(dr.Item(10), String)
                    If Not IsDBNull(dr.Item(11)) Then DataGridView1.Rows.Item(Count).Cells(11).Value = CType(dr.Item(11), String)
                    If Not IsDBNull(dr.Item(12)) Then DataGridView1.Rows.Item(Count).Cells(12).Value = CType(dr.Item(12), String)
                    If Not IsDBNull(dr.Item(13)) Then DataGridView1.Rows.Item(Count).Cells(13).Value = CType(dr.Item(13), String)
                    If Not IsDBNull(dr.Item(14)) Then DataGridView1.Rows.Item(Count).Cells(14).Value = CType(dr.Item(14), String)
                    If Not IsDBNull(dr.Item(15)) Then DataGridView1.Rows.Item(Count).Cells(15).Value = CType(dr.Item(15), String)
                    If Not IsDBNull(dr.Item(16)) Then DataGridView1.Rows.Item(Count).Cells(16).Value = CType(dr.Item(16), String)
                    If Not IsDBNull(dr.Item(17)) Then DataGridView1.Rows.Item(Count).Cells(17).Value = CType(dr.Item(17), String)
                    If Not IsDBNull(dr.Item(18)) Then DataGridView1.Rows.Item(Count).Cells(18).Value = CType(dr.Item(18), String)
                    If Not IsDBNull(dr.Item(19)) Then DataGridView1.Rows.Item(Count).Cells(19).Value = CType(dr.Item(19), String)
                    If Not IsDBNull(dr.Item(20)) Then DataGridView1.Rows.Item(Count).Cells(20).Value = CType(dr.Item(20), String)
                    If Not IsDBNull(dr.Item(21)) Then DataGridView1.Rows.Item(Count).Cells(21).Value = CType(dr.Item(21), String)
                    If Not IsDBNull(dr.Item(22)) Then DataGridView1.Rows.Item(Count).Cells(22).Value = CType(dr.Item(22), String)
                    If Not IsDBNull(dr.Item(23)) Then DataGridView1.Rows.Item(Count).Cells(23).Value = CType(dr.Item(23), String)
                    If Not IsDBNull(dr.Item(24)) Then DataGridView1.Rows.Item(Count).Cells(24).Value = CType(dr.Item(24), String)
                    If Not IsDBNull(dr.Item(25)) Then DataGridView1.Rows.Item(Count).Cells(25).Value = CType(dr.Item(25), String)
                    If Not IsDBNull(dr.Item(26)) Then DataGridView1.Rows.Item(Count).Cells(26).Value = CType(dr.Item(26), String)
                    If Not IsDBNull(dr.Item(27)) Then DataGridView1.Rows.Item(Count).Cells(27).Value = CType(dr.Item(27), String)
                    If Not IsDBNull(dr.Item(28)) Then DataGridView1.Rows.Item(Count).Cells(28).Value = CType(dr.Item(28), String)
                    If Not IsDBNull(dr.Item(29)) Then DataGridView1.Rows.Item(Count).Cells(29).Value = CType(dr.Item(29), String)
                    If Not IsDBNull(dr.Item(30)) Then DataGridView1.Rows.Item(Count).Cells(30).Value = CType(dr.Item(30), String)
                    If Not IsDBNull(dr.Item(31)) Then DataGridView1.Rows.Item(Count).Cells(31).Value = CType(dr.Item(31), String)
                    If Not IsDBNull(dr.Item(32)) Then DataGridView1.Rows.Item(Count).Cells(32).Value = CType(dr.Item(32), String)


                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkSpecs")
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
                    If Not IsDBNull(drLocal.Item(10)) Then DataGridView1.Rows.Item(Count).Cells(10).Value = CType(drLocal.Item(10), String)
                    If Not IsDBNull(drLocal.Item(11)) Then DataGridView1.Rows.Item(Count).Cells(11).Value = CType(drLocal.Item(11), String)
                    If Not IsDBNull(drLocal.Item(12)) Then DataGridView1.Rows.Item(Count).Cells(12).Value = CType(drLocal.Item(12), String)
                    If Not IsDBNull(drLocal.Item(13)) Then DataGridView1.Rows.Item(Count).Cells(13).Value = CType(drLocal.Item(13), String)
                    If Not IsDBNull(drLocal.Item(14)) Then DataGridView1.Rows.Item(Count).Cells(14).Value = CType(drLocal.Item(14), String)
                    If Not IsDBNull(drLocal.Item(15)) Then DataGridView1.Rows.Item(Count).Cells(15).Value = CType(drLocal.Item(15), String)
                    If Not IsDBNull(drLocal.Item(16)) Then DataGridView1.Rows.Item(Count).Cells(16).Value = CType(drLocal.Item(16), String)
                    If Not IsDBNull(drLocal.Item(17)) Then DataGridView1.Rows.Item(Count).Cells(17).Value = CType(drLocal.Item(17), String)
                    If Not IsDBNull(drLocal.Item(18)) Then DataGridView1.Rows.Item(Count).Cells(18).Value = CType(drLocal.Item(18), String)
                    If Not IsDBNull(drLocal.Item(19)) Then DataGridView1.Rows.Item(Count).Cells(19).Value = CType(drLocal.Item(19), String)
                    If Not IsDBNull(drLocal.Item(20)) Then DataGridView1.Rows.Item(Count).Cells(20).Value = CType(drLocal.Item(20), String)
                    If Not IsDBNull(drLocal.Item(21)) Then DataGridView1.Rows.Item(Count).Cells(21).Value = CType(drLocal.Item(21), String)
                    If Not IsDBNull(drLocal.Item(22)) Then DataGridView1.Rows.Item(Count).Cells(22).Value = CType(drLocal.Item(22), String)
                    If Not IsDBNull(drLocal.Item(23)) Then DataGridView1.Rows.Item(Count).Cells(23).Value = CType(drLocal.Item(23), String)
                    If Not IsDBNull(drLocal.Item(24)) Then DataGridView1.Rows.Item(Count).Cells(24).Value = CType(drLocal.Item(24), String)
                    If Not IsDBNull(drLocal.Item(25)) Then DataGridView1.Rows.Item(Count).Cells(25).Value = CType(drLocal.Item(25), String)
                    If Not IsDBNull(drLocal.Item(26)) Then DataGridView1.Rows.Item(Count).Cells(26).Value = CType(drLocal.Item(26), String)
                    If Not IsDBNull(drLocal.Item(27)) Then DataGridView1.Rows.Item(Count).Cells(27).Value = CType(drLocal.Item(27), String)
                    If Not IsDBNull(drLocal.Item(28)) Then DataGridView1.Rows.Item(Count).Cells(28).Value = CType(drLocal.Item(28), String)
                    If Not IsDBNull(drLocal.Item(29)) Then DataGridView1.Rows.Item(Count).Cells(29).Value = CType(drLocal.Item(29), String)
                    If Not IsDBNull(drLocal.Item(30)) Then DataGridView1.Rows.Item(Count).Cells(30).Value = CType(drLocal.Item(30), String)
                    If Not IsDBNull(drLocal.Item(31)) Then DataGridView1.Rows.Item(Count).Cells(31).Value = CType(drLocal.Item(31), String)
                    If Not IsDBNull(drLocal.Item(32)) Then DataGridView1.Rows.Item(Count).Cells(32).Value = CType(drLocal.Item(32), String)

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
            SQLStr = cmbSelect.Text & " from Specifications"
        Else
            SQLStr = cmbSelect.Text & " * from Specifications"
        End If

        If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
        If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
        If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"
        LoadTestData(SQLStr)
DontDoIt:


    End Sub
End Class