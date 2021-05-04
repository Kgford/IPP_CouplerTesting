Module ExcelReports

    Public TOTALSCALE As Single = 10
    Public LOWMIDHIGH As Boolean
    Public ActiveInput As Double
    Public ActiveTest As String = ""
    Private ExcelWkbk As Excel.Workbook
    Private ExcelApp As Excel.Application
    Private Wks As Excel.Worksheet
    Private range As Excel.Range
    Private xlsfilespec As String
    Private xlsfilename As String
    Private sheet As Excel.Worksheet
    Private LocalPath As String
    Private WB_X_SPACE As Double
    Private WB_X_MAX As Double
    Private WB_X_Min As Double
    Public SpurTestNumber As Integer = 0
    Private SubBand As String = "FullBand"
    Public PlotTitle As String = ""
    Public BandState As String = ""
    Public FILENAME As String
    Public NET_FILENAME As String
    Public ItemNum As Integer
   

    Public Sub DataToCells(Cells As String, Data As String)
        sheet.Range(Cells).Value = Data
    End Sub

    Public Sub StartupReport(ExcelTemplatePath As String, ExcelTemplateName As String)
        OpenTemplate(ExcelTemplatePath, ExcelTemplateName)
        xlsfilename = ExcelTemplateName
        SuppressDialog()
        GetObjectReference()

    End Sub
    Private Sub OpenTemplate(ExcelTemplatePath As String, ExcelTemplateName As String)
        xlsfilespec = ExcelTemplatePath & ExcelTemplateName
        ExcelApp = CreateObject("Excel.Application")
        ExcelApp.Visible = True
        ExcelApp.Workbooks.Open(xlsfilespec)
        ExcelApp.Windows(1).WindowState = Excel.XlWindowState.xlMinimized
    End Sub
    Private Sub SuppressDialog()
        ExcelApp.DisplayAlerts = False
    End Sub
   
    Private Sub GetObjectReference()

        ExcelWkbk = ExcelApp.Workbooks(1)
        ExcelWkbk.Windows(1).WindowState = Excel.XlWindowState.xlMaximized
        sheet = ExcelApp.Workbooks(xlsfilename).Worksheets(1)
        sheet.EnableCalculation = True

    End Sub
    Private Sub Print()
        sheet.PrintOut()
    End Sub
    Private Sub CloseReport()

        sheet.Application.Quit()
        '~~> Clean Up
        releaseObject(ExcelApp)
        releaseObject(ExcelWkbk)
        range = Nothing
        sheet = Nothing
        Wks = Nothing
        ExcelWkbk = Nothing
        ExcelApp = Nothing

    End Sub
    '~~> Release the objects
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Public Sub SaveAs(FILENAME As String)
        Try
            ExcelApp.Windows(1).WindowState = Excel.XlWindowState.xlMaximized
            sheet.SaveAs(FILENAME)
        Catch
        End Try
    End Sub
    Public Sub EndReport()
        'Print()
        CloseReport()
    End Sub

    Public Sub SelectSheet(SheetName As String)
        Try
            ExcelApp.Sheets(SheetName).Select()
            ExcelApp.Worksheets(SheetName).Activate()
            sheet = ExcelApp.Workbooks(xlsfilename).ActiveSheet
            System.Threading.Thread.Sleep(30)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LoadChartTitle(SheetName As String, ChartName As String, Title As String)
        Try
            ExcelApp.Sheets(SheetName).Select()
            ExcelApp.ActiveSheet.ChartObjects(ChartName).Activate()
            ExcelApp.ActiveChart.ChartTitle.Select()
            ExcelApp.ActiveChart.ChartTitle.Text = Title
            ExcelApp.Selection.Format.TextFrame2.TextRange.Characters.Text = Title
            System.Threading.Thread.Sleep(30)
        Catch ex As Exception

        End Try
    End Sub

    Public Sub LoadChart1(SheetName As String, ChartTitle As String)

        Try
            Dim StartInLabel As String = ""
            Dim StopInLabel As String = ""
            Dim StartOutLabel As String = ""
            Dim StopOutLabel As String = ""
            Dim MaximumScale As Double
            Dim MinimumScale As Double
            Dim Majorunit As Double
            Dim CrossesAt As Double

            LoadChartTitle(SheetName, "Chart 1", ChartTitle)
            '****** Load and Scale data array **************
            range = sheet.Range("A56:C256")
            range.NumberFormat = "0.0"

            Dim cDataArray(UBound(XArray), 3) As Object
            For r = 0 To UBound(XArray) - 1
                cDataArray(r, 0) = XArray(r)
                cDataArray(r, 1) = YArray(r)
                cDataArray(r, 2) = YArray1(r)
            Next
            SelectSheet(SheetName)
            DataToCells("A54", ChartTitle)
            sheet.Range("A56").Resize(UBound(cDataArray) + 1, 3).Value = cDataArray

            'Y Scale
            MaximumScale = YArray.Max + 1
            MinimumScale = YArray.Min - 1.5

            MaximumScale = CInt(MaximumScale)
            MinimumScale = CInt(MinimumScale)

            Majorunit = Math.Abs((MaximumScale - MinimumScale) / 5)
            ' Majorunit = 0.5
            CrossesAt = YArray.Min - 5

            'Set WB Scaling
            WB_X_Min = MinNoZero(XArray)
            WB_X_MAX = XArray.Max
            WB_X_SPACE = (XArray.Max - MinNoZero(XArray)) / 10

            'scale Y chart axis
            ScaleChart("Chart 1", WB_X_Min, WB_X_MAX, WB_X_SPACE, MinimumScale, MaximumScale, Majorunit, CrossesAt)

        Catch ex As Exception

        End Try
    End Sub
    Public Sub LoadChart2(SheetName As String, ChartTitle As String)

        Dim StartInLabel As String = ""
        Dim StopInLabel As String = ""
        Dim StartOutLabel As String = ""
        Dim StopOutLabel As String = ""
        Dim MaximumScale As Double
        Dim MinimumScale As Double
        Dim Majorunit As Double
        Dim CrossesAt As Double

        LoadChartTitle(SheetName, "Chart 2", ChartTitle)
        '****** Load and Scale data array **************
        range = sheet.Range("D56:E256")
        range.NumberFormat = "0.0"

        Dim cDataArray(UBound(XArray), 2) As Object
        For r = 0 To UBound(XArray) - 1
            cDataArray(r, 0) = XArray(r)
            cDataArray(r, 1) = YArray(r)
        Next
        SelectSheet(SheetName)
        DataToCells("D54", ChartTitle)
        sheet.Range("D56").Resize(UBound(cDataArray) + 1, 2).Value = cDataArray

        'Y Scale
        MaximumScale = YArray.Max + (TOTALSCALE / 2)
        MinimumScale = YArray.Min - (TOTALSCALE / 2)

        MaximumScale = CInt(MaximumScale)
        MinimumScale = CInt(MinimumScale)

        Majorunit = Math.Abs((MaximumScale - MinimumScale) / 5)
        ' Majorunit = SCALESIZE
        CrossesAt = YArray.Min - TOTALSCALE

        'Set WB Scaling
        WB_X_MAX = XArray.Max
        WB_X_SPACE = (XArray.Max - MinNoZero(XArray)) / 10

        'scale Y chart axis

        ScaleChart("Chart 2", WB_X_Min, WB_X_MAX, WB_X_SPACE, MinimumScale, MaximumScale, Majorunit, CrossesAt)


    End Sub

    Public Sub LoadChart3(SheetName As String, ChartTitle As String)

        Dim StartInLabel As String = ""
        Dim StopInLabel As String = ""
        Dim StartOutLabel As String = ""
        Dim StopOutLabel As String = ""
        Dim MaximumScale As Double
        Dim MinimumScale As Double
        Dim Majorunit As Double
        Dim CrossesAt As Double



        LoadChartTitle(SheetName, "Chart 3", ChartTitle)
        '****** Load and Scale data array **************
        range = sheet.Range("F56:G256")
        range.NumberFormat = "0.0"

        Dim cDataArray(UBound(XArray), 2) As Object
        For r = 0 To UBound(XArray) - 1
            cDataArray(r, 0) = XArray(r)
            cDataArray(r, 1) = YArray(r)
        Next
        SelectSheet(SheetName)
        DataToCells("F54", ChartTitle)
        sheet.Range("F56").Resize(UBound(cDataArray) + 1, 2).Value = cDataArray

        'Y Scale
        MaximumScale = YArray.Max + (TOTALSCALE / 2)
        MinimumScale = YArray.Min - (TOTALSCALE / 2)

        MaximumScale = CInt(MaximumScale)
        MinimumScale = CInt(MinimumScale)

        Majorunit = Math.Abs((MaximumScale - MinimumScale) / 5)
        'Majorunit = SCALESIZE
        CrossesAt = YArray.Min - TOTALSCALE

        'Set WB Scaling
        WB_X_Min = MinNoZero(XArray)
        WB_X_MAX = XArray.Max
        WB_X_SPACE = (XArray.Max - MinNoZero(XArray)) / 10

        'scale Y chart axis

        ScaleChart("Chart 3", WB_X_Min, WB_X_MAX, WB_X_SPACE, MinimumScale, MaximumScale, Majorunit, CrossesAt)


    End Sub

    Public Sub LoadChart4(SheetName As String, ChartTitle As String)

        Dim StartInLabel As String = ""
        Dim StopInLabel As String = ""
        Dim StartOutLabel As String = ""
        Dim StopOutLabel As String = ""
        Dim MaximumScale As Double
        Dim MinimumScale As Double
        Dim Majorunit As Double
        Dim CrossesAt As Double

        LoadChartTitle(SheetName, "Chart 4", ChartTitle)
        '****** Load and Scale data array **************
        range = sheet.Range("H56:J256")
        range.NumberFormat = "0.0"

        Dim cDataArray(UBound(XArray), 3) As Object
        For r = 0 To UBound(XArray) - 1
            cDataArray(r, 0) = XArray(r)
            cDataArray(r, 1) = YArray(r)
            cDataArray(r, 2) = YArray1(r)
        Next
        SelectSheet(SheetName)
        DataToCells("H54", ChartTitle)
        sheet.Range("H56").Resize(UBound(cDataArray) + 1, 3).Value = cDataArray

        'Y Scale
        MaximumScale = YArray.Max + 1.5
        MinimumScale = YArray.Min - 1.5

        MaximumScale = CInt(MaximumScale)
        MinimumScale = CInt(MinimumScale)

        Majorunit = Math.Abs((MaximumScale - MinimumScale) / 5)
        'Majorunit = 0.5
        CrossesAt = YArray.Min - 5

        'Set WB Scaling
        WB_X_Min = MinNoZero(XArray)
        WB_X_MAX = XArray.Max
        WB_X_SPACE = (XArray.Max - MinNoZero(XArray)) / 10

        'scale Y chart axis
        ScaleChart("Chart 4", WB_X_Min, WB_X_MAX, WB_X_SPACE, MinimumScale, MaximumScale, Majorunit, CrossesAt)
    End Sub

    Sub ScaleChart(ChartName As String, MinXScale As Double, MaxXScale As Double, XSpace As Double, MinYScale As Double, MaxYScale As Double, YSpace As Double, YCross As Double)
        Try
            ExcelApp.ActiveSheet.ChartObjects(ChartName).Activate()
            ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlCategory).Select()
            ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlCategory).MinimumScale = MinXScale
            ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlCategory).MaximumScale = MaxXScale
            ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlCategory).MajorUnit = XSpace
            ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlCategory).MinorUnit = XSpace
            ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).Select()
            ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).MinimumScale = MinYScale
            ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).MaximumScale = MaxYScale
            ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).MajorUnit = YSpace
            ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).MinorUnit = YSpace
            ExcelApp.ActiveChart.Axes(Excel.XlAxisType.xlValue).CrossesAt = YCross
            '  range("N32").Select()
        Catch ex As Exception

        End Try
    End Sub

End Module
