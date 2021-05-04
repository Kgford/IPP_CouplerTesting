VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XLReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private XLApp As EXCEL.Application
Private XLWB As EXCEL.Workbook
Private XLS As EXCEL.Worksheet
Private ExcelOpen As Boolean
Dim hWndXl As Long
        
' used for database integration
Private Hidden() As String
Private NumHidden As Long
Private XLSFile2Save As String

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub StartExcel()

    Set XLApp = CreateObject("Excel.Application")
    On Error Resume Next
    
    Set XLApp = GetObject(, "Excel.Application")
    '''''''''''''''''''''''''''
    ' If code breaks above, drag current line to second Set XLApp statement below
    ''''''''
    If Err.Number = 429 Then
        Set XLApp = GetObject("", "Excel.Application")
    End If
    '''''''''''''''''''''''''''
    
    XLApp.Caption = "Excel - Autotest"
    hWndXl = FindWindow("XLMAIN", XLApp.Caption)
    XLApp.Caption = Empty 'Set the original caption back
    XLApp.Visible = True
    ExcelOpen = True

End Sub


Public Sub OpenSheet(StartPath As String, Name As String, _
            TemplateName As String, Optional NewTab As Boolean = False, Optional XLSPath As String = "", Optional TemplatePath As String = "")

    Dim i As Long
    Dim Found As Boolean
    ' excel objects...
    Dim Template As String
    'Dim XLSPath As String
    Dim TemplateWB As EXCEL.Workbook
    Dim TemplateWS As EXCEL.Worksheet
    Const MaxFileSize = 0.5 'MB
    Dim File2 As String
    Dim FileCnt As Integer
    Dim NextName As Long
    Dim t As Long
    
    If XLSPath = "" Then
        XLSPath = StartPath
    End If
    
    If TemplatePath = "" Then
        TemplatePath = StartPath & "Templates\"
    End If
    
    ' make sure the xls folder already exists
    If Len(Dir(XLSPath, vbDirectory)) < 1 Then
        MkDir (XLSPath)
    End If
    
    ' create the full filename we will use
    XLSFile2Save = XLSPath & Name & ".xls"
    Template = TemplatePath & TemplateName & ".xls"
        
    StartExcel
    
    ' if our XLS file already exists
NextSheet:
    If Len(Dir(XLSFile2Save)) > 0 Then
        
        'Debug.Print FileSize(XLSFile2Save)
        If Val(FileSize(XLSFile2Save)) > MaxFileSize Then
            FileCnt = FileCnt + 1
            XLSFile2Save = XLSPath & Name & "-" & FileCnt & ".xls"
            GoTo NextSheet
        End If
        
        ' then just open it
        For i = 1 To XLApp.Workbooks.Count
            If XLApp.Workbooks(i).Name = Dir(XLSFile2Save) Then
                Set XLWB = XLApp.Workbooks(i)
                Found = True
            End If
        Next i
        If Not Found Then
            Set XLWB = XLApp.Workbooks.Open(XLSFile2Save)
        End If
        Found = False
        Set TemplateWB = XLApp.Workbooks.Open(Template)
        ' and get the first worksheet
        'Debug.Assert False
        
        For i = 1 To TemplateWB.Worksheets.Count
            Set TemplateWS = TemplateWB.Worksheets(i)
            If UCase(TemplateWS.Name) = "TEMPLATE" Then
                TemplateWS.Copy , XLWB.Worksheets(XLWB.Worksheets.Count)
                Set XLS = XLWB.Worksheets(XLWB.Worksheets.Count)
                NextName = 0
                For t = 1 To XLWB.Worksheets.Count
                    If IsNumeric(XLWB.Worksheets(t).Name) Then
                        If Val(XLWB.Worksheets(t).Name) > NextName Then
                            NextName = XLWB.Worksheets(t).Name
                        End If
                    End If
                Next t
                XLS.Name = NextName + 1
                Found = True
                TemplateWB.Close False
                Exit For
            End If
        Next i
        If Not Found Then
            Err.Raise 123, , "No template sheet found for new tab."
        End If
        
    Else
        
        ' make sure we have a template to copy
        If Len(Dir(Template)) < 1 Then
            Err.Raise 123, , "Missing XLS Template '" & Template & "'"
        End If
        ' open the template
        Set XLWB = XLApp.Workbooks.Open(Template)
        ' save it with our serial nuber
        XLWB.SaveAs XLSFile2Save
        Set XLS = XLWB.Worksheets(XLWB.Worksheets.Count)
        XLS.Name = XLWB.Worksheets.Count - 1
        
    End If

    ' make this the active sheet
    XLS.Activate

End Sub

Public Sub OpenSheetTab(StartPath As String, Name As String, _
            TemplateName As String, TabName As String, Optional XLSPath As String = "")

    Dim i As Long
    Dim Found As Boolean
    Dim Template As String

    If XLSPath = "" Then
        XLSPath = StartPath
    End If
    
    ' make sure the xls folder already exists
    If Len(Dir(XLSPath, vbDirectory)) < 1 Then
        MkDir (XLSPath)
    End If
    
    ' create the full filename we will use
    XLSFile2Save = XLSPath & Name & ".xls"
    Template = StartPath & "Templates\" & TemplateName & ".xls"
        
    StartExcel
    
    ' if our XLS file already exists
    If Len(Dir(XLSFile2Save)) > 0 Then
        ' is it already open in excel?
        For i = 1 To XLApp.Workbooks.Count
            If XLApp.Workbooks(i).Name = Dir(XLSFile2Save) Then
                Set XLWB = XLApp.Workbooks(i)
                Found = True
            End If
        Next i
        ' If not open, the open it
        If Not Found Then
            Set XLWB = XLApp.Workbooks.Open(XLSFile2Save)
        End If
    Else
        ' make sure we have a template to copy
        If Len(Dir(Template)) < 1 Then
            Err.Raise 123, , "Missing XLS Template '" & Template & "'"
        End If
        ' open the template
        Set XLWB = XLApp.Workbooks.Open(Template)
        ' save it with our serial nuber
        XLWB.SaveAs XLSFile2Save
    End If

    Found = False
    For i = 1 To XLWB.Worksheets.Count
        Set XLS = XLWB.Worksheets(i)
        If UCase(XLS.Name) = UCase(TabName) Then
            Found = True
            Exit For
        End If
    Next i
    If Not Found Then
        Err.Raise 123, , "No template sheet found for new tab."
    End If

    ' make this the active sheet
    XLS.Activate

End Sub

'Public Sub OpenSheetOLD(StartPath As String, Name As String, _
'            TemplateName As String, Optional NewTab As Boolean = False)
'
'    Dim i As long
'    Dim Found As Boolean
'    ' excel objects...
'    Dim Template As String
'    Dim XLSPath As String
'    Dim TemplateWB As Excel.Workbook
'
'    XLSPath = StartPath
'
'    ' make sure the xls folder already exists
'    If Len(Dir(XLSPath, vbDirectory)) < 1 Then
'        MkDir (XLSPath)
'    End If
'
'    ' create the full filename we will use
'    XLSFile2Save = XLSPath & Name & ".xls"
'    Template = StartPath & "Templates\" & TemplateName & ".xls"
'
'    ' open excel
'    Set XLApp = New Excel.Application
'    ExcelOpen = True
'
'    ' if our XLS file already exists
'    If Len(Dir(XLSFile2Save)) > 0 Then
'        ' the just open it
'        Set XLWB = XLApp.Workbooks.Open(XLSFile2Save)
'        ' if we did not already create an XLS for this serial
'    Else
'        ' make sure we have a template to copy
'        If Len(Dir(Template)) < 1 Then
'            Err.Raise 123, , "Missing XLS Template '" & Template & "'"
'        End If
'        ' open the template
'        Set XLWB = XLApp.Workbooks.Open(Template)
'        ' save it with our serial nuber
'        XLWB.SaveAs XLSFile2Save
'    End If
'
'    If NewTab Then
'        ' and get the first worksheet
'        For i = 1 To XLWB.Worksheets.Count
'            Set XLS = XLWB.Worksheets(i)
'            If UCase(XLS.Name) = "TEMPLATE" Then
'                XLS.Copy XLWB.Worksheets(XLWB.Worksheets.Count)
'                Set XLS = XLWB.Worksheets(XLWB.Worksheets.Count - 1)
'                XLS.Name = XLWB.Worksheets.Count - 1
'                Found = True
'                Exit For
'            End If
'        Next i
'        If Not Found Then
'            Err.Raise 123, , "No template sheet found for new tab."
'        End If
'        'Count = ExcelWorkBook.Worksheets.Count
'        'Set ExcelSheet = ExcelWorkBook.Worksheets.Add
'    Else
'        Set XLS = XLWB.Worksheets(1) 'ExcelWorkBook.Worksheets.Count)
'    End If
'    ' show excel
'    XLApp.Visible = True
'
'    ' make this the active sheet
'    XLS.Activate
'
'End Sub
'
'

Public Sub PrintSheet(Optional FromPage As Long = -1, Optional ToPage As Long = -1)

    If FromPage > -1 And ToPage > -1 And FromPage <= ToPage Then
        XLS.PrintOut FromPage, ToPage
    Else
        XLS.PrintOut
    End If

End Sub

Public Sub Save()

    If Len(Dir(XLSFile2Save)) < 1 Then
        XLWB.SaveAs XLSFile2Save, , , , , , , xlLocalSessionChanges
    Else
        XLWB.Save
    End If

End Sub


Public Sub SaveAs(Dir As String, File As String)

    XLSFile2Save = Dir & File
    XLWB.SaveAs XLSFile2Save, , , , , , , xlLocalSessionChanges
  

End Sub


Public Sub OpenTemplate(Template As String, Optional NewTab As Boolean = False)
        
        
    Dim i As Long
    Dim Found As Boolean
        
    StartExcel
    
    ' make sure we have a template to copy
    If Len(Dir(Template)) < 1 Then
        Err.Raise 123, , "Missing XLS Template '" & Template & "'"
    End If
    ' open the template
    Set XLWB = XLApp.Workbooks.Open(Template)
    
    If NewTab Then
        ' and get the first worksheet
        For i = 1 To XLWB.Worksheets.Count
            Set XLS = XLWB.Worksheets(i)
            If UCase(XLS.Name) = "TEMPLATE" Then
                XLWB.Worksheets(1).Copy XLWB.Worksheets(XLWB.Worksheets.Count)
                Set XLS = XLWB.Worksheets(XLWB.Worksheets.Count - 1)
                XLS.Name = XLWB.Worksheets.Count - 1
                Found = True
                Exit For
            End If
        Next i
        If Not Found Then
            Err.Raise 123, , "No template sheet found for new tab."
        End If
        'Count = ExcelWorkBook.Worksheets.Count
        'Set ExcelSheet = ExcelWorkBook.Worksheets.Add
    Else
        Set XLS = XLWB.Worksheets(1) 'ExcelWorkBook.Worksheets.Count)
    End If
    ' show excel
    XLApp.Visible = True

    ' make this the active sheet
    XLS.Activate

End Sub

Public Sub OpenNewSheet()

    ' excel objects...
    Dim ExcelWorkBook As EXCEL.Workbook
    Dim ExcelOpen As Boolean
        
    ' open excel
    StartExcel
    ' add our work book
    Set ExcelWorkBook = XLApp.Workbooks.Add
    ' store pointer to active sheet
    Set XLS = ExcelWorkBook.Worksheets(1)
    ' show excel
    XLApp.Visible = True
    ' make this the active sheet
    XLS.Activate
End Sub

Public Sub LoadImage(cell As String, File As String)

    XLS.Application.Visible = True
    SetForegroundWindow hWndXl
    XLS.Range(cell).Activate
    XLS.Activate
    ' use the insert picture method in excel
    SendKeys "%IPF", True
    SendKeys File, True
    SendKeys "%R", True
    
    ' now shrink the image so it is the same size as the graphs
    'Delay 5000
    'SendKeys "%OI1", True
    'SendKeys "{LEFT}%H90{ENTER}", True
    ''''''''''''''''''''''''''''''''''''''''''''

End Sub

'Public Sub SendTraceToExcel(Row As Long, Col As Long, T1 As trc.XYTrace)
'
'    Dim i As Long
'    Dim CurRow As Long
'    Dim Str As String
'
'    CurRow = Row
'
'    For i = 0 To T1.Count - 1
'        Str = Str & T1.XData(i) & Chr(9) & T1.YData(i) & Chr(13)
'        'XLS.Cells(CurRow, Col) = T1.XData(i)
'        'XLS.Cells(CurRow, Col + 1) = T1.YData(i)
'        'CurRow = CurRow + 1
'    Next i
'
'    Clipboard.Clear
'    Clipboard.SetText Str
'
'    XLS.Range(RowColToCell(Row, Col)).Activate
'    XLS.Paste
'
'End Sub

Public Sub CopyToCell(cell As String)

    XLS.Activate
    XLS.Range(cell).Activate
    XLS.Paste
    
End Sub

Public Sub CopyToRowCol(Row As Long, Col As Long)

    Dim cell As String

    cell = RowColToCell(Row, Col)

    XLS.Activate
    XLS.Range(cell).Activate
    XLS.Paste
    
End Sub

Public Function RowColToCell(Row As Long, Col As Long) As String

    Dim C1 As Long
    Dim C2 As Long
    Dim cell As String
    
    C1 = Col Mod 26
    C2 = Int(Col / 26)
    
    If C2 > 0 Then
        cell = Chr(C2 - 1 + Asc("A"))
    End If
        
    cell = cell & Chr(C1 + Asc("A") - 1) & Row
    RowColToCell = cell
    
End Function

Public Sub CellToRowCol(ByRef Row As Long, ByRef Col As Long, ByRef cell As String)

    Dim LenCell As Long
    Dim i As Long
    Dim ASCC As Long

    ' parse the cell into a row and column...
    LenCell = Len(cell)
    For i = 1 To LenCell
        ASCC = Asc(Mid(cell, i, 1))
        If ASCC > Asc("0") And ASCC <= Asc("9") Then
            ' move forward an extra cell to the data column (graph col specified above)
            Col = Asc(UCase(Left(cell, i - 1))) - Asc("A") + 1
            Row = CInt(Val(Right(cell, LenCell - i + 1)))
            Exit For
        End If
    Next i

End Sub

Public Sub LoadADORst(Rst As ADODB.Recordset, cell As String)

    Dim UsedFields() As Long
    Dim NumUsedFields As Long
    Dim i As Long
    Dim j As Long
    Dim ThisField As String
    Dim KeyIndex As Long
    Dim Hide As Boolean
    Dim itmText As String
    Dim wid As Long
    Dim DefSize As Long
    Dim Key As String
    Dim Row As Long
    Dim Col As Long
    
    CellToRowCol Row, Col, cell
    
    ' detect unwanted fields here and create array of wanted field indexes
    For i = 0 To Rst.Fields.Count - 1
        ' retreive current field name
        ThisField = Rst.Fields(i).Name
        
        ' initialize flag
        Hide = False
        ' check each hidden field
        For j = 0 To NumHidden - 1
            ' do we want to show this field
            If Hidden(j) = ThisField Then
                Hide = True
            End If
        Next j
        ' if the field not in the hidden list, add it
        If Not Hide Then
            ' increment number of use fields
            NumUsedFields = NumUsedFields + 1
            ' make space in used fields array
            ReDim Preserve UsedFields(NumUsedFields) As Long
            'save the index
            UsedFields(NumUsedFields - 1) = i
        End If
    Next i
    
    For i = 0 To NumUsedFields - 1
        ' add a column header for each field
        XLS.Cells(Row, Col + i) = Rst.Fields(UsedFields(i)).Name
    Next i
    Row = Row + 1
    ' do not do any work if there are no records
    If Rst.EOF And Rst.BOF Then
        Exit Sub
    End If
    
    ' shorten calls to recordset
    With Rst
        
        ' go to the first record
        .MoveFirst
        
        ' and loop though every record
        While Not .EOF
            ' add all the sub items info
            For i = 0 To NumUsedFields - 1
                XLS.Cells(Row, Col + i) = .Fields(UsedFields(i)) & ""
            Next i
            Row = Row + 1
            ' next record
            .MoveNext
        Wend
        
    End With
    
    XLS.Range(cell, Chr(Col + NumUsedFields + Asc("A") - 2) & Row).Activate
    XLS.Activate
    SendKeys "%OCA", True
    DoEvents
    
End Sub

Public Sub HideField(FieldName As String)

    NumHidden = NumHidden + 1
    
    ReDim Preserve Hidden(NumHidden) As String
    Hidden(NumHidden - 1) = FieldName
    
End Sub



Public Function Sheet() As EXCEL.Worksheet

    Set Sheet = XLS

End Function

Public Sub WriteToCell(cell As String, Text As String)

    Dim Row As Long
    Dim Col As Long
    
    CellToRowCol Row, Col, cell

    XLS.Cells(Row, Col) = Text
    
End Sub

Public Sub WriteToRowCol(Row As Long, Col As Long, Text As String, Optional BoldFont As Boolean = False)

    XLS.Cells(Row, Col) = Text

    If BoldFont Then
        XLS.Cells(Row, Col).Select
        ActiveCell.Font.Bold = True
        XLS.Cells(1, 1).Select
    End If
End Sub

Public Sub Show()

    SetForegroundWindow hWndXl

End Sub

Public Sub OpenSheetNoTemplateNoNewTab(StartPath As String, Name As String)

    Dim i As Long
    Dim Found As Boolean
    Const MaxFileSize = 0.5 'MB
    Dim File2 As String
    Dim FileCnt As Integer
    Dim XLSPath As String
    
'    If XLSPath = "" Then
        XLSPath = StartPath
'    End If
    
    ' make sure the xls folder already exists
    If Len(Dir(XLSPath, vbDirectory)) < 1 Then
        MkDir (XLSPath)
    End If
    
    ' create the full filename we will use
    XLSFile2Save = XLSPath & Name & ".xls"
        
    StartExcel
    
    ' if our XLS file already exists
NextSheet:
    If Len(Dir(XLSFile2Save)) > 0 Then
        
        'Debug.Print FileSize(XLSFile2Save)
        If Val(FileSize(XLSFile2Save)) > MaxFileSize Then
            FileCnt = FileCnt + 1
            XLSFile2Save = XLSPath & Name & "-" & FileCnt & ".xls"
            GoTo NextSheet
        End If
        
        ' then just open it
        For i = 1 To XLApp.Workbooks.Count
            If XLApp.Workbooks(i).Name = Dir(XLSFile2Save) Then
                Set XLWB = XLApp.Workbooks(i)
                Set XLS = XLWB.Worksheets(XLWB.Worksheets.Count)
                XLS.Name = XLWB.Worksheets.Count - 1
                Found = True
            End If
        Next i
        If Not Found Then
            Set XLWB = XLApp.Workbooks.Open(XLSFile2Save)
        End If
        Found = False
        
    End If

    Set XLS = XLWB.ActiveSheet

    ' make this the active sheet
    'XLS.Activate

End Sub

Public Function GetExcelVersion() As Long
    On Error Resume Next
    
    Set XLApp = GetObject(, "Excel.Application")
    '''''''''''''''''''''''''''
    ' If code breaks above, drag current line to second Set XLApp statement below
    ''''''''
    If Err.Number = 429 Then
        Set XLApp = GetObject("", "Excel.Application")
    End If
    GetExcelVersion = Val(XLApp.Application.Version)
    Set XLApp = Nothing
End Function

Public Sub AddCloneSheet(Template As String, NewSheet As String)
Dim num As Long
Dim Length As Long
Dim checkNum As String
On Error GoTo AppendName

num = 0
XLApp.Sheets(Template).Select
XLApp.Sheets(Template).Copy Before:=Sheets(2)

Repeat:
    XLApp.ActiveSheet.Name = NewSheet
    XLApp.ActiveSheet.SetBackgroundPicture FileName:="" ' Remove Template Background from new sheet
    Set XLS = XLApp.ActiveSheet
    XLS.Activate
    Exit Sub

AppendName:
    For Length = 1 To Int(Len(NewSheet))
       checkNum = Left(NewSheet, Length)
       If IsNumeric(Mid(checkNum, Length, 1)) Then
           num = num + 1
           NewSheet = NewSheet & "_" & num
           Exit For
       End If
    Next Length
 GoTo Repeat
    
End Sub




Public Sub HideColumn(Column As String)
   
    XLApp.Columns(Column & ":" & Column).Select
    XLApp.Selection.EntireColumn.Hidden = True
    
End Sub
