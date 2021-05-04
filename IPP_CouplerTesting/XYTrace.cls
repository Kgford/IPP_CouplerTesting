VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Trace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Function OpenTrace() As Long
Dim SQLstr As String
Dim Title As String
Dim ATS As ADODB.Recordset
On Error GoTo Trap
        
    
    Title = " New Title"
    SQLstr = "SELECT * from Trace where JobNumber = '" & frmAUTOTEST.cmbJob.Text & "' And Title = '" & Title & "'"
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If ATS.EOF Then
         SQLstr = "Insert Into Trace (JobNumber, Title) values ('" & frmAUTOTEST.cmbJob.Text & "','" & Title & "')"
        NetConn.Execute SQLstr
        Delay 100
        
        SQLstr = "SELECT * from Trace where JobNumber = '" & frmAUTOTEST.cmbJob.Text & "' And Title = '" & Title & "'"
        Set ATS = New ADODB.Recordset
        ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    End If
    
    OpenTrace = ATS!ID
    ATS!Title = Null
    ATS.Update
    ATS.Close
    
    
Exit Function
Trap:
End Function


Public Function AddPoint(TraceID As Long, Index As Long, XData As Double, YData As Double)
Dim SQLstr As String
Dim ATS As ADODB.Recordset
On Error GoTo Trap
    
    SQLstr = "Insert Into TracePoints (TraceID,XData,YData,Idx) values ('" & TraceID & "','" & XData & "','" & YData & "','" & Index & "')"
    NetConn.Execute SQLstr
    
Exit Function
Trap:
    
End Function
Public Function AddAllPoints(TraceID As Long, Pts As Long, XData() As Double, YData() As Double)
Dim SQLstr As String
Dim ATS As ADODB.Recordset
Dim i As Integer
On Error GoTo Trap

i = 0
Do While i < Pts
    SQLstr = "Insert Into TracePoints (TraceID,XData,YData,Idx) values ('" & TraceID & "','" & XData(i) & "','" & YData(i) & "','" & i & "')"
    NetConn.Execute SQLstr
    i = i + 1
Loop
 
Exit Function
Trap:
    
End Function

Public Function AddCoeffCalChar(Name As String, Coeff As String)
Dim SQLstr As String
Dim ATS As ADODB.Recordset

On Error GoTo Trap
    
    SQLstr = "SELECT * from  CoeffCal where StringName =  '" & Name & "'"
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If ATS.EOF Then
        SQLstr = "Insert Into CoeffCal (StringName) values ('" & Name & "')"
        NetConn.Execute SQLstr
        
        SQLstr = "SELECT * from  CoeffCal where StringName =  '" & Name & "'"
        Set ATS = New ADODB.Recordset
        ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    End If
        ATS!StringName = Name
        ATS!String = Coeff
        ATS.Update
        ATS.Close
    
      
        
    
Exit Function
Trap:
    
End Function


Public Function SaveTrace(TraceID As Long, Optional ShowPlot As Boolean = False)
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    
    ATS!ProgVer = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    ATS!ProgTitle = App.Title
    ATS!ActiveDate = Now
    ATS!InstrumentCalDue = GetInstrCalDue
    ATS.Update
    ATS.Close
Exit Function
Trap:
    
End Function


Public Function SerialNumber(TraceID As Long, TraceSerialNumber As String)
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    
    ATS!SerialNumber = TraceSerialNumber
    ATS.Update
    ATS.Close

    Exit Function
Trap:
    
End Function


Public Function Title(TraceID As Long, TraceTitle As String)
     Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    
    ATS!Title = TraceTitle
    ATS.Update
    ATS.Close

    Exit Function
Trap:
    
End Function


Public Function CalDate(TraceID As Long, Calibration As Date)
     Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    
    ATS!CalibrationDate = Calibration
    ATS.Update
    ATS.Close

    Exit Function
Trap:
    
End Function

Public Function Notes(TraceID As Long, TestNotes As String)
     Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    
    ATS!Notes = TestNotes
    ATS.Update
    ATS.Close

    Exit Function
Trap:
    
End Function

Public Function Points(TraceID As Long, TracePoints As Integer)
     Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    
    ATS!Points = TracePoints
    ATS.Update
    ATS.Close

    Exit Function
Trap:
    
End Function

Public Function SpecID(TraceID As Long, SpecificationID As Long)
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    
    ATS!SpecID = SpecificationID
    ATS.Update
    ATS.Close

    Exit Function
Trap:
    
End Function

Public Function TestID(TraceID As Long, TestIDs As Long)
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    
    ATS!TestID = TestIDs
    ATS.Update
    ATS.Close

    Exit Function
Trap:
    
End Function

Public Function Workstation(TraceID As Long, ComputerName As String)
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    
    ATS!Workstation = ComputerName
    ATS.Update
    ATS.Close

    Exit Function
Trap:
    
End Function

Public Function GetTraceIDByTitle(Title As String, Optional Serial As String = "", Optional Job As String = "") As Long
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    If Serial = "" Then Serial = frmAUTOTEST.UUTLabel.Caption & frmAUTOTEST.UUTCount.Text
    If Job = "" Then Job = frmAUTOTEST.cmbJob.Text
    
    SQLstr = "SELECT * from Trace where JobNumber =  '" & Job & "' and SerialNumber = '" & Serial & "' and Title = '" & Title & "'"
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
        
        If ATS.EOF Then
            GetTraceIDByTitle = -1
        Else
            GetTraceIDByTitle = ATS!ID
        End If

    ATS.Close

    Exit Function
Trap:
    
End Function

Public Function GetTraceIDByTitleLocal(Title As String, Optional Serial As String = "", Optional Job As String = "") As Long
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    If Serial = "" Then Serial = frmAUTOTEST.UUTLabel.Caption & frmAUTOTEST.UUTCount.Text
    If Job = "" Then Job = frmAUTOTEST.cmbJob.Text
    
    SQLstr = "SELECT * from Trace where JobNumber =  '" & Job & "' and SerialNumber = '" & Serial & "' and Title = '" & Title & "'"
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
        
        If ATS.EOF Then
            GetTraceIDByTitleLocal = -1
        Else
            GetTraceIDByTitleLocal = ATS!ID
        End If

    ATS.Close

    Exit Function
Trap:
    
End Function
Public Function GetTraceIDByTitleNet(Title As String, Optional Serial As String = "", Optional Job As String = "") As Long
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    If Serial = "" Then Serial = frmAUTOTEST.UUTLabel.Caption & frmAUTOTEST.UUTCount.Text
    If Job = "" Then Job = frmAUTOTEST.cmbJob.Text
    
    SQLstr = "SELECT * from Trace where JobNumber =  '" & Job & "' and SerialNumber = '" & Serial & "' and Title = '" & Title & "'"
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
        
        If ATS.EOF Then
            GetTraceIDByTitleNet = -1
        Else
            GetTraceIDByTitleNet = ATS!ID
        End If

    ATS.Close

    Exit Function
Trap:
    
End Function

Public Function MinMax(TraceID As Long, Optional retMin As Double, Optional retMax As Double, Optional retMinX As Long, Optional retMaxX As Long)
    Dim i As Integer
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap

    SQLstr = "SELECT * from TracePoints where TraceID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    
    i = 0
    While Not ATS.EOF
        If i = 0 Then
            retMax = ATS!YData
            retMin = ATS!YData
        Else
            If ATS!YData > retMax Then
                retMax = ATS!YData
                retMaxX = i
             End If
            If ATS!YData < retMin Then
                retMin = ATS!YData
                retMainX = i
            End If
        End If
        i = i + 1
        ATS.MoveNext
     Wend
   
Exit Function
Trap:
    
End Function


Public Function GetYData(TraceID As Long, Index As Integer) As Double
    Dim i As Integer
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap

    SQLstr = "SELECT * from TracePoints where TraceID =  " & TraceID & " and Idx = " & Index & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If ATS.EOF Then
        GetYData = 0
    Else
        GetYData = ATS!YData
   End If
Exit Function
Trap:
    
End Function
Public Function GetYDataLocal(TraceID As Long, Index As Integer) As Double
    Dim i As Integer
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap

'    SQLstr = "SELECT * from TracePoints where TraceID =  " & TraceID & " and Idx = " & Index & ""
'    Set ATS = New ADODB.Recordset
'    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
'    If ATS.EOF Then
'        GetYDataLocal = 0
'    Else
'        GetYDataLocal = ATS!YData
'   End If
 GetYDataLocal = YData
Exit Function
Trap:
    
End Function

Public Function GetYDataNet(TraceID As Long, Index As Integer) As Double
    Dim i As Integer
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap

'    SQLstr = "SELECT * from TracePoints where TraceID =  " & TraceID & " and Idx = " & Index & ""
'    Set ATS = New ADODB.Recordset
'    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
'    If ATS.EOF Then
'        GetYDataNet = 0
'    Else
'        GetYDataNet = ATS!YData
'   End If
   GetYDataNet = YData
Exit Function
Trap:
    
End Function


Public Function GetXData(TraceID As Long, Index As Integer) As Double
    Dim i As Integer
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap

'    SQLstr = "SELECT * from TracePoints where TraceID =  " & TraceID & " and Idx = " & Index & ""
'    Set ATS = New ADODB.Recordset
'    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
'    If ATS.EOF Then
'        GetXData = 0
'        YData = 0
'    Else
'        GetXData = ATS!XData
'        YData = ATS!YData
'    End If
    GetXData = YData
Exit Function
Trap:
    
End Function
Public Function GetXDataLocal(TraceID As Long, Index As Integer) As Double
    Dim i As Integer
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap

    SQLstr = "SELECT * from TracePoints where TraceID =  " & TraceID & " and Idx = " & Index & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If ATS.EOF Then
        GetXDataLocal = 0
        YData = 0
    Else
        GetXDataLocal = ATS!XData
        YData = ATS!YData
    End If
   
Exit Function
Trap:
    
End Function

Public Function GetXDataNet(TraceID As Long, Index As Integer) As Double
    Dim i As Integer
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap

    SQLstr = "SELECT * from TracePoints where TraceID =  " & TraceID & " and Idx = " & Index & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If ATS.EOF Then
        GetXDataNet = 0
        YData = 0
    Else
        GetXDataNet = ATS!XData
        YData = ATS!YData
    End If
   
Exit Function
Trap:
    
End Function



Public Function GetPoints(TraceID As Long) As Double
    Dim i As Integer
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap

    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If ATS.EOF Then
        GetPoints = 0
    Else
        GetPoints = ATS!Points
    End If
   
Exit Function
Trap:
    
End Function
Public Function GetPointsLocal(TraceID As Long) As Double
    Dim i As Integer
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap

    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If ATS.EOF Then
        GetPointsLocal = 0
    Else
        GetPointsLocal = ATS!Points
    End If
   
Exit Function
Trap:
    
End Function

Public Function GetPointsNet(TraceID As Long) As Double
    Dim i As Integer
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap

    SQLstr = "SELECT * from Trace where ID =  " & TraceID & ""
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If ATS.EOF Then
        GetPointsNet = 0
    Else
        GetPointsNet = ATS!Points
    End If
   
Exit Function
Trap:
    
End Function


Public Function GetInstrCalDue() As Date
 Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    On Error GoTo Trap
    
    SQLstr = "SELECT * from InstCalDue where WorkStation =  '" & GetComputerName & "'" 'Get Workstation code
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If Not ATS.EOF Then GetInstrCalDue = ATS!InstrumentCalDue
    
Exit Function
Trap:
End Function

Public Function GetCalCoeffs()
    Dim i As Integer
    Dim SQLstr As String
    Dim ATS As ADODB.Recordset
    Dim Title As String
    Dim TraceID As Long
    Dim p, Channel, SwPos As Integer
    
    CalCoeff1 = ""
    CalCoeff2 = ""
    CalCoeff3 = ""
    
    Channel = 2
    SwPos = 1
    Title = "CalCoeff CH" & Channel & " SwPos: " & SwPos
    SQLstr = "SELECT * from Trace where Title =  '" & Title & "'"
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If Not ATS.EOF Then
        TraceID = ATS!TraceID
        p = ATS!Points
        ATS.Close
        SQLstr = "SELECT * from TracePoints where TraceID =  " & TraceID & ""
        Set ATS = New ADODB.Recordset
        ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
        While Not ATS.EOF
            CalCoeff1 = CalCoeff1 & ATS!CoeffCal
            ATS.MoveNext
        Wend
        ATS.Close
    End If
    
    Channel = 2
    SwPos = 2
    Title = "CalCoeff CH" & Channel & " SwPos: " & SwPos
    SQLstr = "SELECT * from Trace where Title =  '" & Title & "'"
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If Not ATS.EOF Then
        TraceID = ATS!TraceID
        p = ATS!Points
        ATS.Close
        SQLstr = "SELECT * from TracePoints where TraceID =  " & TraceID & ""
        Set ATS = New ADODB.Recordset
        ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
        While Not ATS.EOF
            CalCoeff1 = CalCoeff2 & ATS!CoeffCal
            ATS.MoveNext
        Wend
        ATS.Close
    End If
    
    Channel = 2
    SwPos = 3
    Title = "CalCoeff CH" & Channel & " SwPos: " & SwPos
    SQLstr = "SELECT * from Trace where Title =  '" & Title & "'"
    Set ATS = New ADODB.Recordset
    ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
    If Not ATS.EOF Then
        TraceID = ATS!TraceID
        p = ATS!Points
        ATS.Close
        SQLstr = "SELECT * from TracePoints where TraceID =  " & TraceID & ""
        Set ATS = New ADODB.Recordset
        ATS.Open SQLstr, NetConn, adOpenDynamic, adLockOptimistic
        While Not ATS.EOF
            CalCoeff3 = CalCoeff1 & ATS!CoeffCal
            ATS.MoveNext
        Wend
        ATS.Close
    End If
    

End Function








