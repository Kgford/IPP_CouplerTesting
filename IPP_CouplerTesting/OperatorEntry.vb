Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection

Public Class OperatorEntry

    Public Sub New()

        Try
            Dim SQLStr As String

            InitializeComponent()
            If frmAUTOTEST.DeleteOp.Checked Then
                DeleteOperator.Visible = True
            Else
                DeleteOperator.Visible = False
            End If
            Me.cmbOperator.Items.Clear()
            Me.cmbOperator.Items.Add("New Operator")
            SQLStr = "SELECT DISTINCT Operator from Effeciency where Operator is not null and  Operator <> ''"
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    Me.cmbOperator.Items.Add(dr.Item(0))
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
                    Me.cmbOperator.Items.Add(drLocal.Item(0))
                End While
                atsLocal.Close()
            End If
            If frmAUTOTEST.ckROBOT.Checked Then
                TestRunningSignal(True) ' Note False/False tells the Robot if test is running

            End If

        Catch
            If Not GetLastUser() = Nothing Then cmbOperator.Text = GetLastUser()
        End Try
    End Sub


    Private Sub btOK_Click(sender As Object, e As EventArgs) Handles btOK.Click
        If txtOperator.Text = "" Then
            lblOperator.Visible = True
            lblOperator.ForeColor = Color.Red
            lblOperator.Text = "*Enter a Valid Name"
        Else
            User = txtOperator.Text
            SQL.SaveUser(txtOperator.Text)
            Me.Close()
        End If
    End Sub

    Private Sub cmbOperator_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbOperator.SelectedIndexChanged
        If Me.cmbOperator.GetItemText(Me.cmbOperator.SelectedItem) = "New Operator" Then
            lblOperator.Visible = True
            txtOperator.Text = ""
        Else
            lblOperator.Visible = False
            txtOperator.Text = Me.cmbOperator.GetItemText(Me.cmbOperator.SelectedItem)
        End If
    End Sub

    Private Sub DeleteOperator_Click(sender As Object, e As EventArgs) Handles DeleteOperator.Click
        If txtOperator.Text = "" Then
            MsgBox("Please Choose Operator to Delete", , "No Operator choosen!")
            Me.Close()
            Exit Sub
        End If
        If MsgBox("Are you sure you want to erase" & txtOperator.Text & "From the records", vbYesNo, "Are you sure??") = vbNo Then
            Me.Close()
            Exit Sub
        End If
        SQL.DeleteOperator(txtOperator.Text)
        Me.Close()
    End Sub
End Class