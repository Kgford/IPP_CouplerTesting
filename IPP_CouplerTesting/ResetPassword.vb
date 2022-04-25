Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions

Public Class ResetPassword
    ' Replace [A-Z] with \p{Lu}, to allow for Unicode uppercase letters.
    Private upper As New System.Text.RegularExpressions.Regex("[A-Z]")
    Private lower As New System.Text.RegularExpressions.Regex("[a-z]")
    Private number As New System.Text.RegularExpressions.Regex("[0-9]")
    ' Special is "none of the above".
    Private special As New System.Text.RegularExpressions.Regex("[^a-zA-Z0-9]")
    Public Sub New()
        MyBase.New()
        InitializeComponent()

    End Sub


    Function ValidatePassword(ByVal pwd As String, Optional ByVal minLength As Integer = 8, Optional ByVal numUpper As Integer = 1, Optional ByVal numLower As Integer = 2, Optional ByVal numNumbers As Integer = 2, Optional ByVal numSpecial As Integer = 2) As String
        Dim Str1 As String = ""
        Dim Str2 As String = ""

        ' Check the length.
        If Len(pwd) < minLength Then Str1 = "Invalid"

        ' Check for minimum number of occurrences.
        If upper.Matches(pwd).Count < numUpper Then
            Str2 = "Fail Upper"
            GoTo SendOut
        End If

        If lower.Matches(pwd).Count < numLower Then
            Str2 = "Fail Lower"
            GoTo SendOut
        End If
        If number.Matches(pwd).Count < numNumbers Then
            Str2 = "Fail Number"
            GoTo SendOut
        End If
        If special.Matches(pwd).Count < numSpecial Then Str2 = "Strong"

SendOut:
        ' Passed all checks.
        Return Str1 & " " & Str2
    End Function



    Private Sub txtNewPassword2_TextChanged(sender As Object, e As EventArgs) Handles txtNewPassword2.TextChanged
        If txtNewPassword2.Text = txtNewPassword1.Text Then
            lblMatch.Text = True
        Else
            lblMatch.Text = False
        End If
        If lblMatch.Text = True And lblStrong.Text.ToUpper.Contains("STRONG") Then
            btUpdate.Visible = True
        End If

    End Sub

    Private Sub txtNewPassword1_TextChanged(sender As Object, e As EventArgs) Handles txtNewPassword1.TextChanged
        lblStrong.Text = ValidatePassword(txtNewPassword1.Text)
        If lblMatch.Text = True And lblStrong.Text.ToUpper.Contains("STRONG") Then
            btUpdate.Visible = True
        End If
    End Sub

    Private Sub btUpdate_Click(sender As Object, e As EventArgs) Handles btUpdate.Click
        Dim ActivePassword As String
        If txtPassword.Text = SupervisorPassword Then
            If lblMatch.Text = False Then
                MYMsgBox("Passwords must Match. Please redo", , "Re-enter Passwords")
                txtResponse.Text = "Passwords must Match. Please redo"
                txtPassword.Text = ""
                txtNewPassword1.Text = ""
                txtNewPassword2.Text = ""
                Me.Refresh()
                Exit Sub
            End If
            If Not lblStrong.Text.ToUpper.Contains("STRONG") Then
                MYMsgBox("Password is not strong. Please redo", , "Re-enter Passwords")
                txtResponse.Text = "Password is not strong. Please redo"
                txtPassword.Text = ""
                txtPassword.Text = ""
                txtNewPassword1.Text = ""
                txtNewPassword2.Text = ""
                Me.Refresh()
                Exit Sub
            End If
            ActivePassword = txtNewPassword1.Text
            SavePassword(ActivePassword)
            txtNewPassword1.Text = ""
            txtNewPassword2.Text = ""
            txtPassword.Text = ""
            txtResponse.Text = "Password changed"
            Me.Refresh()
            Delay(1000)
            txtResponse.Text = ""
            lblStrong.Text = "EXIT"
            lblMatch.Text = "EXIT"
            Delay(500)
            Me.Refresh()
            Me.Close()
        Else
            txtPassword.Text = ""
            txtNewPassword1.Text = ""
            txtNewPassword2.Text = ""
            txtResponse.Text = "Enter the Supervisor Passwod"
            MYMsgBox("Enter the Supervisor Passwor", MsgBoxStyle.Critical, "Password Required")
            txtResponse.Text = ""
            Me.Refresh()
        End If

    End Sub

    
End Class