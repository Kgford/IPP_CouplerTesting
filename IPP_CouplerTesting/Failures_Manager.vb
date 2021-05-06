Public Class Failures_Manager
    Private test1 As String
    Private test2 As String
    Private test3 As String
    Private test4 As String
    Private test5 As String

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()

        If SpecType = "90 DEGREE COUPLER" Or SpecType.Contains("BALUN") Or SpecType = "COMBINER/DIVIDER" Then
            test1 = "Insertion Loss"
            test2 = "Return Loss"
            test3 = "Isolation"
            test4 = "Amplitude balance"
            test5 = "Phase Balance"
        ElseIf InStr(SpecType, "DIRECTIONAL COUPLER") Then
            test1 = "Insertion Loss"
            test2 = "Return Loss"
            test3 = "Coupling"
            test4 = "Directivity"
            test5 = "Coupled Balance"
        End If

        If Test1Failing Then
            Top_label.Text = test1 & " Failures"
        ElseIf Test2Failing Then
            Top_label.Text = test2 & " Failures"
        ElseIf TEST3Failing Then
            Top_label.Text = test3 & " Failures"
        ElseIf TEST4Failing Then
            Top_label.Text = test4 & " Failures"
        ElseIf TEST5Failing Then
            Top_label.Text = test5 & " Failures"
        Else
            Top_label.Text = "Global Failures Manager"
        End If

    End Sub

    Private Sub btAgain_Click(sender As Object, e As EventArgs) Handles btAgain.Click
        If txtPassword.Text = SupervisorPassword Then
            Me.Close()
        Else
            MsgBox("Enter the Supervisor Password", MsgBoxStyle.Critical, "Password Required")
        End If
    End Sub

    Private Sub btIgnore_Click(sender As Object, e As EventArgs) Handles btIgnore.Click
        If txtPassword.Text = SupervisorPassword Then
            If Test1Failing Then
                Test1Fail_bypass = True
            ElseIf Test2Failing Then
                Test2Fail_bypass = True
            ElseIf TEST3Failing Then
                TEST3Fail_bypass = True
            ElseIf TEST4Failing Then
                TEST4Fail_bypass = True
            ElseIf TEST5Failing Then
                TEST5Fail_bypass = True
            Else
                GlobalFail_bypass = True
            End If
            Me.Close()
        Else
            MsgBox("Enter the Supervisor Password", MsgBoxStyle.Critical, "Password Required")
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If txtPassword.Text = SupervisorPassword Then
            If CheckBox1.CheckState = CheckState.Checked Then
                Master_bypass = True
                SQL.SaveBypass(1)
            Else
                Master_bypass = False
                SQL.SaveBypass(0)
            End If
            Me.Close()
        Else
            MsgBox("Enter the Supervisor Password", MsgBoxStyle.Critical, "Password Required")
        End If
    End Sub
End Class