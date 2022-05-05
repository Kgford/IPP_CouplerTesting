Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection

Public Class AddArtWorkRevision
    Private Sub txtArtwork_TextChanged(sender As Object, e As EventArgs) Handles txtArtwork.TextChanged
        UUTReset = True
        txtArtwork.SelectionStart = Len(txtArtwork.Text)
        txtArtwork.Text = Trim(txtArtwork.Text.ToUpper)
        Artwork = txtArtwork.Text
    End Sub
    Private Sub txtRev_TextChanged(sender As Object, e As EventArgs) Handles txtRev.TextChanged
        UUTReset = True
        Rev = txtRev.Text
    End Sub
    Private Sub txtPanel_TextChanged(sender As Object, e As EventArgs) Handles txtPanel.TextChanged
        UUTReset = True
        Panel = txtPanel.Text
        If txtRev.Text.Length() = 1 Then
            If txtRev.Text.Contains("*") Then
                txtRev.Text = "*" + txtPanel.Text
            Else
                txtRev.Text = "0" + txtRev.Text
            End If
        End If
        Rev = txtRev.Text
    End Sub
    Private Sub txtsector_TextChanged(sender As Object, e As EventArgs) Handles txtSector.TextChanged
        UUTReset = True
        txtSector.SelectionStart = Len(txtSector.Text)
        txtSector.Text = Trim(txtSector.Text.ToUpper)
        Sector = txtSector.Text
        If txtRev.Text.Length() = 1 Then
            If txtRev.Text.Contains("*") Then
                txtRev.Text = "*" + txtPanel.Text
            Else
                txtRev.Text = "0" + txtRev.Text
            End If
        End If
        Rev = txtRev.Text
        If txtPanel.Text.Length() = 1 Then
            If txtPanel.Text.Contains("*") Then
                txtPanel.Text = "*" + txtPanel.Text
            Else
                txtPanel.Text = "0" + txtPanel.Text
            End If
        End If
        Panel = txtPanel.Text
    End Sub
    Private Sub txtLOT_TextChanged(sender As Object, e As EventArgs) Handles txtLOT.TextChanged
        UUTReset = True
        If txtRev.Text.Length() = 1 Then
            If txtRev.Text.Contains("*") Then
                txtRev.Text = "*" + txtPanel.Text
            Else
                txtRev.Text = "0" + txtRev.Text
            End If
        End If
        Rev = txtRev.Text
        If txtPanel.Text.Length() = 1 Then
            If txtPanel.Text.Contains("*") Then
                txtPanel.Text = "*" + txtPanel.Text
            Else
                txtPanel.Text = "0" + txtPanel.Text
            End If
        End If
        Panel = txtPanel.Text
        LOT = txtLOT.Text
    End Sub


    Private Sub btrev_Click(sender As Object, e As EventArgs) Handles btrev.Click

        If txtLOT.Text = "" And txtLOT.Text = "" And txtPanel.Text = "" And txtSector.Text = "" Then
            MYMsgBox("Please enter all data")
            Exit Sub
        End If
        If txtRev.Text.Length() = 1 Then
            If txtRev.Text.Contains("*") Then
                txtRev.Text = "*" + txtPanel.Text
            Else
                txtRev.Text = "0" + txtRev.Text
            End If
        End If
        Rev = txtRev.Text
        If txtPanel.Text.Length() = 1 Then
            If txtPanel.Text.Contains("*") Then
                txtPanel.Text = "*" + txtPanel.Text
            Else
                txtPanel.Text = "0" + txtPanel.Text
            End If
        End If
        Panel = txtPanel.Text
        If txtLOT.Text.Length() = 1 Then
            txtLOT.Text = "000000000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 2 Then
            txtLOT.Text = "00000000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 3 Then
            txtLOT.Text = "0000000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 4 Then
            txtLOT.Text = "000000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 5 Then
            txtLOT.Text = "00000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 6 Then
            txtLOT.Text = "0000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 7 Then
            txtLOT.Text = "000000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 8 Then
            txtLOT.Text = "00000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 9 Then
            txtLOT.Text = "0000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 10 Then
            txtLOT.Text = "000" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 11 Then
            txtLOT.Text = "00" + txtLOT.Text
        ElseIf txtLOT.Text.Length() = 12 Then
            txtLOT.Text = "0" + txtLOT.Text
        End If

        ArtworkRevision = Artwork + Rev + Panel + Sector + LOT
        UUTReset = True
        Me.Close()
    End Sub

   
    

    

  
   
End Class