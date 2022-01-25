Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection

Public Class AddArtWorkRevision

    Private Sub btrev_Click(sender As Object, e As EventArgs) Handles btrev.Click
        ArtworkRevision = txtRevision.Text
        UUTReset = True
        Me.Close()
    End Sub

    Private Sub txtRevision_TextChanged(sender As Object, e As EventArgs) Handles txtRevision.TextChanged
        txtRevision.SelectionStart = Len(txtRevision.Text)
        txtRevision.Text = Trim(txtRevision.Text.ToUpper)
    End Sub
End Class