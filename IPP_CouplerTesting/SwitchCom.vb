Module SwitchCom
    '****************************************************************************************************************************************
    '*************************************Created  by Michael Ford  of Automated Test Solutions 05/06/14*************************************
    '*************************************Driver Software for Mini Circuits RF Switch  USB-1SP4T-A18*****************************************
    '****************************************************************************************************************************************

    Private MyPTE1 As New mcl_RF_Switch_Controller64.USB_RF_SwitchBox
    ' Instantiate new switch object, assign to MyPTE1
    Private MyPTE2 As New mcl_RF_Switch_Controller64.USB_RF_SwitchBox
    ' Instantiate new switch object, assign to MyPTE2

    Public Function Connect() As Integer
        Connect = MyPTE1.Connect
    End Function

    Public Sub Disconnect()
        MyPTE1.Disconnect()
    End Sub


    Public Function GetSNlist() As String
        Dim SN As String = ""
        Dim NumberOfSwitches As Integer

        NumberOfSwitches = MyPTE1.Get_Available_SN_List(SN)
        GetSNlist = SN
    End Function

    Public Function GetNumberOfSwitches() As String
        Dim SN As String = ""
        GetNumberOfSwitches = MyPTE1.Get_Available_SN_List(SN)
    End Function

    Public Function Get24VConnection() As String
        If MyPTE1.Get_24V_Indicator > 0 Then
            Get24VConnection = "24V supply connected"
        Else
            Get24VConnection = "24V supply Not connected"
        End If

    End Function

    Public Function SetSwitchPosition(Position As Byte) As Integer

        SetSwitchPosition = MyPTE1.Set_SP4T_COM_To(Position)
    End Function

    Public Function GetSwitchPosition(StatusRet As Integer) As Integer
        GetSwitchPosition = MyPTE1.GetSwitchesStatus(StatusRet)
    End Function


    Public Function GetFirmware() As Integer
        GetFirmware = MyPTE1.GetFirmware
    End Function


    Public Sub TestSwitches()
        Dim ConnectionGood As Integer
        Dim SerialList As String
        Dim SwitchNumber As Integer
        Dim VoltageConnected As String
        Dim status As String
        Dim StatusRet As Integer
        Dim Firmware As Integer



        SwitchNumber = GetNumberOfSwitches()
        SerialList = GetSNlist()
        Firmware = GetFirmware()


        ConnectionGood = Connect() ' Note requires a few milliseconds to connect
        System.Threading.Thread.Sleep(2000)

        VoltageConnected = Get24VConnection()

        status = SetSwitchPosition(1) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        status = GetSwitchPosition(StatusRet) ' note Status Return in Binary

        status = SetSwitchPosition(2) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        status = GetSwitchPosition(StatusRet) ' note Status Return in Binary

        status = SetSwitchPosition(3) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        status = GetSwitchPosition(StatusRet) ' note Status Return in Binary

        status = SetSwitchPosition(4) 'note: Status 0 = Error,Status 1 = Switched, Status 1 = Switch commmand recieved, no 24V
        status = GetSwitchPosition(StatusRet) ' note Status Return in Binary
        Disconnect()

    End Sub

End Module
