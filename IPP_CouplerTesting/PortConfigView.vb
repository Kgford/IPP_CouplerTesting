Imports System.Data
Imports System.Data.SqlClient

Public Class PortConfigView

    Public Sub New()
        MyBase.New()
        Dim SQLStr As String
        Try
            InitializeComponent()

            SQLStr = "Select * from PortConfig"
            LoadTestData(SQLStr)

            cmbSelect.Text = "Select"

        Catch
        End Try
    End Sub
    Private Sub LoadTestData(SQLStr As String)
        Dim NameStr(31) As String

        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        NameStr(0) = "ID"
        NameStr(1) = "JobNumber"
        NameStr(2) = "PartNumber"
        NameStr(3) = "J1J1"
        NameStr(4) = "J1J2"
        NameStr(5) = "J1J3"
        NameStr(6) = "J1J4"
        NameStr(7) = "J1J5"
        NameStr(8) = "J2J1"
        NameStr(9) = "J2J2"
        NameStr(10) = "J2J3"
        NameStr(11) = "J2J4"
        NameStr(12) = "J3J1"
        NameStr(13) = "J3J2"
        NameStr(14) = "J3J3"
        NameStr(15) = "J3J4"
        NameStr(16) = "J4J1"
        NameStr(17) = "J4J2"
        NameStr(18) = "J4J3"
        NameStr(19) = "J4J4"
        Try

            For x = 0 To 31
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
            SQLStr = cmbSelect.Text & " from PortConfig"
        Else
            SQLStr = cmbSelect.Text & " * from PortConfig"
        End If

        If Not cmbPARM1.Text = Nothing Then SQLStr = SQLStr & " where " & cmbPARM1.Text & " = '" & txtPARM1.Text & "'"
        If Not cmbPARM2.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM2.Text & " = '" & txtPARM2.Text & "'"
        If Not cmbPARM3.Text = Nothing Then SQLStr = SQLStr & " and " & cmbPARM3.Text & " = '" & txtPARM3.Text & "'"
        LoadTestData(SQLStr)
DontDoIt:


    End Sub

End Class