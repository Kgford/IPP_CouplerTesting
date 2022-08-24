Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection

Public Class GoldenList
    Dim intActivate As Integer = 0
    Dim WishedForCell As DataGridViewCell = Nothing

    Public Sub New()
        MyBase.New()
        Dim SQLStr As String
        Try
            InitializeComponent()
            SQLStr = "Select * from TestData "
            If Not Job = Nothing Then
                SQLStr = "SELECT * from TestFixtures where PartNumber = '" & Part & "'"

            End If

            LoadTestData(SQLStr)

        Catch
        End Try
    End Sub

    Private Sub LoadTestData(SQLStr As String)
        Dim NameStr(5) As String

        DataGridView1.Rows.Clear()
        DataGridView1.Columns.Clear()
        NameStr(0) = "ID"
        NameStr(1) = "PartNumber"
        NameStr(2) = "Fixture"
        NameStr(3) = "Plunger"
        NameStr(4) = "GoldRev"
        NameStr(5) = "FixNum."


        Try

            For x = 0 To 5
                Dim col As New DataGridViewTextBoxColumn

                col.DataPropertyName = NameStr(x)
                col.HeaderText = NameStr(x)
                col.Name = NameStr(x)
                DataGridView1.Columns.Add(col)
            Next

            If CheckforRow(SQLStr, "NetworkData") > 0 Then GoTo GotIt

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
                End While
                ats.Close()
            Else
                Dim strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDatabaseFolder("NetworkData")
                Dim atsLocal As New OleDb.OleDbConnection
                atsLocal.ConnectionString = strConnectionString
                Dim cmd As OleDb.OleDbCommand = New OleDb.OleDbCommand(SQLStr, atsLocal)
                atsLocal.Open()
                System.Threading.Thread.Sleep(10)
                Dim drLocal As OleDb.OleDbDataReader = cmd.ExecuteReader
                While Not drLocal.Read = Nothing
                    Dim Count As Integer = DataGridView1.Rows.Add()

                    If Not IsDBNull(drLocal.Item(1)) Then DataGridView1.Rows.Item(Count).Cells(1).Value = CType(drLocal.Item(1), String)
                    If Not IsDBNull(drLocal.Item(2)) Then DataGridView1.Rows.Item(Count).Cells(2).Value = CType(drLocal.Item(2), String)
                    If Not IsDBNull(drLocal.Item(3)) Then DataGridView1.Rows.Item(Count).Cells(3).Value = CType(drLocal.Item(3), String)
                    If Not IsDBNull(drLocal.Item(4)) Then DataGridView1.Rows.Item(Count).Cells(4).Value = CType(drLocal.Item(4), String)
                    If Not IsDBNull(drLocal.Item(5)) Then DataGridView1.Rows.Item(Count).Cells(5).Value = CType(drLocal.Item(5), String)
                End While
                atsLocal.Close()
            End If
            '*******************************************************************************
        Catch
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        GoldenPN = DataGridView1.Item(1, e.RowIndex).Value.ToString
        Goldenfixture = DataGridView1.Item(2, e.RowIndex).Value.ToString
        GoldenPlunger = DataGridView1.Item(3, e.RowIndex).Value.ToString
        GoldenRev = DataGridView1.Item(4, e.RowIndex).Value.ToString
        Goldenfixture_num = DataGridView1.Item(5, e.RowIndex).Value.ToString
        Me.Close()

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

        GoldenPN = DataGridView1.Item(1, e.RowIndex).Value.ToString
        Goldenfixture = DataGridView1.Item(2, e.RowIndex).Value.ToString
        GoldenPlunger = DataGridView1.Item(3, e.RowIndex).Value.ToString
        GoldenRev = DataGridView1.Item(4, e.RowIndex).Value.ToString
        Goldenfixture_num = DataGridView1.Item(5, e.RowIndex).Value.ToString
        Me.Close()

    End Sub
    

End Class