Imports System.Runtime.InteropServices
Public Class CalGizmo
    Private loaded As Boolean = False
    Private iniPathName As String = "C:/GIZMO/code/myconfig.ini"
    Private p0_j0, p0_j1, p0_j2, p0_j3, p0_j4 As String
    Private p1_j0, p1_j1, p1_j2, p1_j3, p1_j4 As String
    Private px_j0, px_j1, px_j2, px_j3, px_j4 As String
    Private p0x_j0, p0x_j1, p0x_j2, p0x_j3, p0x_j4 As String
    Private p1x_j0, p1_j1x, p1x_j2, p1x_j3, p1x_j4 As String
    Private GripGrab, txGripPlace As String


    Private Sub Button2_Click(sender As Object, e As EventArgs)
        If Not loaded Then
            LoadminiLAB()
            loaded = True
        End If
        TestCompleteSignal(False)
        tx_p0_j0.Text = GetReadySignal()
    End Sub

    Private Sub btSave_p0_Click(sender As Object, e As EventArgs) Handles btSave_p0.Click
        saveConfigurationVal(iniPathName, "J0_P0", tx_p0_j0.Text)
        saveConfigurationVal(iniPathName, "J1_P0", tx_p0_j1.Text)
        saveConfigurationVal(iniPathName, "J2_P0", tx_p0_j2.Text)
        saveConfigurationVal(iniPathName, "J3_P0", tx_p0_j3.Text)
        saveConfigurationVal(iniPathName, "J4_P0", tx_p0_j4.Text)
    End Sub

    Private Sub btSave_p1_Click(sender As Object, e As EventArgs) Handles btSave_p1.Click
        saveConfigurationVal(iniPathName, "J0_P1", tx_p1_j0.Text)
        saveConfigurationVal(iniPathName, "J1_P1", tx_p1_j1.Text)
        saveConfigurationVal(iniPathName, "J2_P1", tx_p1_j2.Text)
        saveConfigurationVal(iniPathName, "J3_P1", tx_p1_j3.Text)
        saveConfigurationVal(iniPathName, "J4_P1", tx_p1_j4.Text)
    End Sub

    Private Sub btSave_px_Click(sender As Object, e As EventArgs) Handles btSave_px.Click
        saveConfigurationVal(iniPathName, "J0_P0_New", tx_px_j0.Text)
        saveConfigurationVal(iniPathName, "J1_P0_New", tx_px_j1.Text)
        saveConfigurationVal(iniPathName, "J2_P0_New", tx_px_j2.Text)
        saveConfigurationVal(iniPathName, "J3_P0_New", tx_px_j3.Text)
        saveConfigurationVal(iniPathName, "J4_P0_New", tx_px_j4.Text)
    End Sub

    Private Sub btSave_px0_Click(sender As Object, e As EventArgs) Handles btSave_px0.Click
        saveConfigurationVal(iniPathName, "J0_P0_Pass", tx_px0_j0.Text)
        saveConfigurationVal(iniPathName, "J1_P0_Pass", tx_px0_j1.Text)
        saveConfigurationVal(iniPathName, "J2_P0_Pass", tx_px0_j2.Text)
        saveConfigurationVal(iniPathName, "J3_P0_Pass", tx_px0_j3.Text)
        saveConfigurationVal(iniPathName, "J4_P0_Pass", tx_px0_j4.Text)
    End Sub

    Private Sub btSave_p1x_Click(sender As Object, e As EventArgs) Handles btSave_p1x.Click
        saveConfigurationVal(iniPathName, "J0_P0_Fail", tx_px1_j0.Text)
        saveConfigurationVal(iniPathName, "J1_P0_Fail", tx_px1_j1.Text)
        saveConfigurationVal(iniPathName, "J2_P0_Fail", tx_px1_j2.Text)
        saveConfigurationVal(iniPathName, "J3_P0_Fail", tx_px1_j3.Text)
        saveConfigurationVal(iniPathName, "J4_P0_Fail", tx_px1_j4.Text)
    End Sub

    Private Sub btSave_GP_Click(sender As Object, e As EventArgs) Handles btSave_GP.Click
        saveConfigurationVal(iniPathName, "gripper_place", tx_GripPlace.Text)
    End Sub

    Private Sub btSave_GG_Click(sender As Object, e As EventArgs) Handles btSave_GG.Click
        saveConfigurationVal(iniPathName, "gripper_grab", tx_GripGrab.Text)
    End Sub

    Private Sub btDin_Click(sender As Object, e As EventArgs) Handles btDin.Click
        Try
            Dim Bits(5) As Integer
            tx_b0.Text = ""
            tx_b1.Text = ""
            tx_b2.Text = ""
            tx_b3.Text = ""
            tx_b4.Text = ""
            Bits = GetInBits()
            tx_b0.Text = Bits(0)
            tx_b1.Text = Bits(1)
            tx_b2.Text = Bits(2)
            tx_b3.Text = Bits(3)
            tx_b4.Text = Bits(4)
        Catch ex As Exception
            MsgBox("Input BIT Error", , "Please check your MINILab")
        End Try

    End Sub

  
    Private Sub btSend_Click_1(sender As Object, e As EventArgs) Handles btSend.Click
        If txCH0.Text > 0 Then
            TestRunningSignal(True)
        Else
            TestRunningSignal(False)
        End If

        If txCH1.Text > 0 Then
            TestCompleteSignal(True)
        Else
            TestCompleteSignal(False)
        End If

        

    End Sub

   
End Class