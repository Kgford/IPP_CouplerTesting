Public Class Supervisor

    Private Sub btAgain_Click(sender As Object, e As EventArgs) Handles btAgain.Click
        If txtPassword.Text = SupervisorPassword Then
            Master_bypass = True
            Percent_bypass = True
        Else
            MsgBox("Enter the Supervisor Password", MsgBoxStyle.Critical, "Password Required")
            Master_bypass = False
            Percent_bypass = False
        End If
        Me.Close()
    End Sub

End Class