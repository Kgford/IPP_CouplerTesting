Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection

Public Class AddArtWorkRevision

    Private Sub btrev_Click(sender As Object, e As EventArgs) Handles btrev.Click
        ArtworkRevision = txtArtwork.Text
        
        If txtLOT.Text <> "" Then
            LOT = txtLOT.Text
        End If
        If txtPanel.Text <> "" Then
            Panel = txtPanel.Text
        End If
        If txtQuadrant.Text <> "" Then
            Quadrant = txtQuadrant.Text
        End If
        UUTReset = True
        Me.Close()
    End Sub
    Private Sub txtArtwork_TextChanged(sender As Object, e As EventArgs) Handles txtArtwork.TextChanged
        UUTReset = True
        txtArtwork.SelectionStart = Len(txtArtwork.Text)
        txtArtwork.Text = Trim(txtArtwork.Text.ToUpper)
        Artwork = txtArtwork.Text

        ArtworkRevision = txtArtwork.Text
         If txtLOT.Text <> "" Then
            LOT = txtLOT.Text
        End If
        If txtPanel.Text <> "" Then
            Panel = txtPanel.Text
        End If
        If txtQuadrant.Text <> "" Then
            Quadrant = txtQuadrant.Text
        End If

    End Sub

    Private Sub txtPanel_TextChanged(sender As Object, e As EventArgs) Handles txtPanel.TextChanged
        UUTReset = True
        txtPanel.SelectionStart = Len(txtPanel.Text)
        txtPanel.Text = Trim(txtPanel.Text.ToUpper)
        ArtworkRevision = txtArtwork.Text
         If txtLOT.Text <> "" Then
            LOT = txtLOT.Text
        End If
        If txtPanel.Text <> "" Then
            Panel = txtPanel.Text
        End If
        If txtQuadrant.Text <> "" Then
            Quadrant = txtQuadrant.Text
        End If
    End Sub
    Private Sub txtQuadrant_TextChanged(sender As Object, e As EventArgs) Handles txtQuadrant.TextChanged
        UUTReset = True
        txtQuadrant.SelectionStart = Len(txtQuadrant.Text)
        txtQuadrant.Text = Trim(txtQuadrant.Text.ToUpper)
        ArtworkRevision = txtArtwork.Text
        If txtLOT.Text <> "" Then
            LOT = txtLOT.Text
        End If
        If txtPanel.Text <> "" Then
            Panel = txtPanel.Text
        End If
        If txtQuadrant.Text <> "" Then
            Quadrant = txtQuadrant.Text
        End If
    End Sub

    Private Sub txtLOT_TextChanged(sender As Object, e As EventArgs) Handles txtLOT.TextChanged
        UUTReset = True
        txtQuadrant.SelectionStart = Len(txtQuadrant.Text)
        txtQuadrant.Text = Trim(txtQuadrant.Text.ToUpper)
        ArtworkRevision = txtArtwork.Text
         If txtLOT.Text <> "" Then
            LOT = txtLOT.Text
        End If
        If txtPanel.Text <> "" Then
            Panel = txtPanel.Text
        End If
        If txtQuadrant.Text <> "" Then
            Quadrant = txtQuadrant.Text
        End If
    End Sub
End Class