

Public Class DirectIO
    Private ioInstr As Ivi.Visa.Interop.FormattedIO488
    Private ioGPIB As Ivi.Visa.Interop.IGpib
    'Public timeout As Integer
    Public TerminationCharacterEnabled As Boolean
    Public TerminationCharacter As Char

    Sub New(ByVal address As String, ByVal locked As Boolean, ByVal timeout As Integer)
        Dim mgr As Ivi.Visa.Interop.ResourceManager
        Try
            If Len(address) > 6 Then
                mgr = New Ivi.Visa.Interop.ResourceManager
                ioInstr = New Ivi.Visa.Interop.FormattedIO488
                ioInstr.IO() = mgr.Open(address)
                ioInstr.IO.Timeout() = timeout
            Else
                mgr = New Ivi.Visa.Interop.ResourceManager
                ioInstr = New Ivi.Visa.Interop.FormattedIO488
                ioInstr.IO() = mgr.Open(address)
                ioInstr.IO.Timeout() = timeout
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "IO Error")
        End Try
    End Sub


    Public Function Timeout(ByVal time As Integer) As Boolean
        ioInstr.IO.Timeout() = time
    End Function

    Public Function GetTimeout() As Integer
        GetTimeout = ioInstr.IO.Timeout
    End Function

    Public Function disconnect() As Boolean
        Try
            ioInstr.IO.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "IO Error")
        End Try

    End Function

    Public Function ExecuteReset() As Boolean
        Try
            ioInstr.FlushWrite
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "IO Error")
        End Try

    End Function
    Public Function ExecuteClear() As Boolean
        Try
            ioInstr.IO.Clear()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "IO Error")
        End Try

    End Function

    Public Function WriteLine(ByVal command As String) As Boolean
        Try
            ioInstr.WriteString(command)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "IO Error")
        End Try
    End Function

    Public Function WriteNumber(ByVal command As Double) As Boolean
        Try
            ioInstr.WriteNumber(command)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "IO Error")
        End Try
    End Function

    Public Function Read() As String
        Dim result As String
        Dim i As Integer
        Try
            result = ioInstr.ReadString
            i = result.Length
            result = result.Remove(i - 1, 1)
            Return result
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "IO Error")
            Return "Fail"
        End Try
    End Function

    Public Function ReadStatusByte() As Integer
        Dim mySTB As Integer
        mySTB = ioInstr.IO.ReadSTB()
        Return mySTB
    End Function

    Public Function GPIBLocal() As Boolean
        ioGPIB.Close()
    End Function
    Public Function ExecuteLocal() As Boolean
        ioInstr.IO.UnlockRsrc()
    End Function

    Public Function UnbufferedRead(ByVal length As Integer) As Byte()
        Dim myByte As Byte()
        myByte = ioInstr.IO.Read(length)
        Return myByte
    End Function

    Public Function WriteByte(ByVal msg As Integer) As Boolean
        Dim myByte As Byte()

        myByte = BitConverter.GetBytes(msg)
        ioInstr.IO.Write(myByte, UBound(myByte) + 1)
    End Function

    Public Function UnbufferedWrite(ByVal msg As Byte()) As Boolean
        'Dim myByte As Byte()
        ioInstr.IO.Write(msg, UBound(msg) + 1)
    End Function

    Public Function Clear() As Boolean
        ioInstr.IO.Clear()
    End Function

    Public Function ReadIeeeBlockAsDoubleArray() As Double()
        ReadIeeeBlockAsDoubleArray = ioInstr.ReadIEEEBlock(Ivi.Visa.Interop.IEEEASCIIType.ASCIIType_Any, False, True)
    End Function

    Public Function ReadListAsDoubleArray() As Double()
        Dim myAry As Double()
        Dim myStr As String
        Dim xStr() As String
        Dim x As Integer
        myStr = ioInstr.ReadString
        xStr = Split(myStr, ",")
        For x = 0 To UBound(xStr)
            ReDim Preserve myAry(x)
            myAry(x) = Val(xStr(x))
        Next
        Return myAry

    End Function

    Public Function ReadListAs2DDoubleArray() As Double(,)
        Try
            Dim myAry As Double(,)
            Dim myAry1 As Double()
            Dim myAry2 As Double()
            Dim myStr As String
            Dim xStr() As String
            Dim x As Integer
            Dim y As Integer = 0
            System.Threading.Thread.Sleep(1000)
            myStr = ioInstr.ReadString
            xStr = Split(myStr, ",")
            For x = 0 To UBound(xStr) - 1 Step 2
                ReDim Preserve myAry(1, y)
                ReDim Preserve myAry1(y)
                ReDim Preserve myAry2(y)
                If Val(xStr(x)) = 0 Then GoTo Skip
                myAry1(y) = Val(xStr(x))
                myAry2(y) = Val(xStr(x + 1))
                myAry(0, y) = myAry1(y)
                myAry(1, y) = myAry2(y)
                y = y + 1
Skip:
            Next
            ReadListAs2DDoubleArray = myAry
        Catch ex As Exception
        End Try

    End Function

    Public Function ReadIeeeBlockAsInt32Array() As Int32()
        ReadIeeeBlockAsInt32Array = ioInstr.ReadIEEEBlock(Ivi.Visa.Interop.IEEEBinaryType.BinaryType_I4, False, True)
    End Function

    Public Function ReadNumberAsInt16() As Int16
        Dim myint As Int16
        'Not needed yet..... 
        'Return myint
    End Function
End Class