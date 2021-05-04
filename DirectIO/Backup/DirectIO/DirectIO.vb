

Public Class DirectIO
    Private ioInstr As Ivi.Visa.Interop.FormattedIO488
    Private ioGPIB As Ivi.Visa.Interop.IGpib
    Public timeout As Integer
    Public TerminationCharacterEnabled As Boolean
    Public TerminationCharacter As Char


    Sub New(ByVal address As String, ByVal locked As Boolean, ByVal timeout As Integer)
        Dim mgr As Ivi.Visa.Interop.ResourceManager
        Try
            If Len(address) > 6 Then
                mgr = New Ivi.Visa.Interop.ResourceManager
                ioInstr = New Ivi.Visa.Interop.FormattedIO488
                ioInstr.IO() = mgr.Open(address)
            Else
                mgr = New Ivi.Visa.Interop.ResourceManager
                ioInstr = New Ivi.Visa.Interop.FormattedIO488
                ioInstr.IO() = mgr.Open(address)
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "IO Error")

        End Try


    End Sub

    Public Function disconnect() As Boolean
        Try
            ioInstr.IO.Close()
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

    'Public Function Timeout(ByVal time As Integer)


    'End Function

    Public Function GPIBLocal() As Boolean
        'Place holder
    End Function

    Public Function UnbufferedRead(ByVal length As Integer) As Byte()
        Dim myByte As Byte()
        myByte = ioInstr.IO.Read(length)
        Return myByte
    End Function
    Public Function UnbufferedWrite(ByVal msg As Byte()) As Boolean
        'Dim myByte As Byte()
        ioInstr.IO.Write(msg, UBound(msg) + 1)

    End Function
    Public Function Clear() As Boolean

        ioInstr.IO.Clear()

    End Function

    Public Function ReadIeeeBlockAsDoubleArray() As Double()

        Dim myByte As Byte()
        'Place holder. Need to complete when needed
        'Return myByte
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

    Public Function ReadIeeeBlockAsInt32Array() As Int32()
        ReadIeeeBlockAsInt32Array = ioInstr.ReadIEEEBlock(Ivi.Visa.Interop.IEEEBinaryType.BinaryType_I4, False, True)
    End Function

    Public Function ReadNumberAsInt16() As Int16
        Dim myint As Int16
        'Not needed yet..... 
        'Return myint
    End Function
End Class