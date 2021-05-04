Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection

Public Class AddArtWorkRevision

    Private Sub btrev_Click(sender As Object, e As EventArgs) Handles btrev.Click
        ArtworkRevision = txtRevision.Text
        Me.Close()
    End Sub
End Class