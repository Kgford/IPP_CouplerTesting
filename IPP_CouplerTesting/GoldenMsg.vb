Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection



Public Class GoldenMsg
    Private SpecIndex As Integer
    Private SpecTest1 As String
    Private SpecTest2 As String
    Private SpecTest3 As String
    Private SpecTest4 As String
    Private SpecTest4_exp As String
    Private SpecTest5 As String
    Private Test1_limit As Double
    Private Test2_limit As Double
    Private Test3_limit As Double
    Private Test4_limit As Double
    Private Test5_limit As Double
    Private ThisData1_old As String = ""
    Private ThisData1H_old As String = ""
    Private ThisData1L_old As String = ""
    Private ThisData2_old As String = ""
    Private ThisData3_old As String = ""
    Private ThisData3H_old As String = ""
    Private ThisData3L_old As String = ""
    Private ThisData4_old As String = ""
    Private ThisData4L_old As String = ""
    Private ThisData4H_old As String = ""
    Private ThisData5_old As String = ""
    Private ThisData5L_old As String = ""
    Private ThisData5H_old As String = ""
    Private offset1 As String
    Private offset2 As String
    Private offset3 As String
    Private offset4 As String
    Private offset5 As String
    Private FixtureSaved As Boolean = True

    Public Sub New()
        InitializeComponent()
        Dim DeltaPass As Boolean = True
        Dim TraceDeltaPass As Boolean = True
        GoldenRev = "N/A"
        Try
            TestFixture()
            LoadSpecs()
            TestData()
            If GoldenHistory Then
                lblOlddata.Text = "Data from " & GoldenDate
            Else
                lblOlddata.Text = "No Data Available"
                btOK.Visible = False
            End If
            loadpanel()
            If Not IsDBNull(GetSpecification("Offset1")) Then offset1 = GetSpecification("Offset1")
            If Not IsDBNull(GetSpecification("Offset2")) Then offset2 = GetSpecification("Offset2")
            If Not IsDBNull(GetSpecification("Offset3")) Then offset3 = GetSpecification("Offset3")
            If Not IsDBNull(GetSpecification("Offset4")) Then offset4 = GetSpecification("Offset4")
            If Not IsDBNull(GetSpecification("Offset5")) Then offset5 = GetSpecification("Offset5")
            Test1()
            Test2()
            Test3()
            Test4()
            Test5()
            If GoldenHistory Then ' There is gold unit data in the database
                If GoldenRunComplete Then ' Recent test completed
                    If PF1.Text > Test1_limit Or PF2.Text > Test2_limit Or PF3.Text > Test3_limit Or PF4.Text > Test4_limit Or PF5.Text > Test5_limit Then DeltaPass = False
                    If PFTrace1.Text > Test1_limit Or PFTrace2.Text > Test2_limit Or PFTrace3.Text > Test3_limit Or PFTrace4.Text > Test4_limit Or PFTrace5.Text > Test5_limit Then TraceDeltaPass = False
                    If GoldenRunPass And DeltaPass And TraceDeltaPass Then
                        GoldenMode = False
                        btTest.Visible = False
                        btBypass.Visible = False
                        Label6.Visible = False
                        Password.Visible = False
                        SpecPF.BackColor = Color.Gold
                        DeltaPF.BackColor = Color.Gold
                        TracePF.BackColor = Color.Gold
                    ElseIf GoldenRunPass And Not DeltaPass And Not TraceDeltaPass Then
                        txtprompt.Text = "The Golden Unit Passes, but it has changed. Contact your supervisor"
                        GoldenMode = True
                        btOK.Visible = False
                        btBypass.Visible = True
                        Label6.Visible = True
                        Password.Visible = True
                        SpecPF.BackColor = Color.Gold
                        DeltaPF.BackColor = Color.Red
                        DeltaPF.ForeColor = Color.Gold
                        DeltaPF.Text = "FAILED"
                        TracePF.BackColor = Color.Red
                        TracePF.Text = "FAILED"
                    ElseIf GoldenRunPass And Not DeltaPass Then
                        txtprompt.Text = "The Golden Unit Passes, but data has changed. Contact your supervisor"
                        GoldenMode = True
                        btOK.Visible = False
                        btBypass.Visible = True
                        Label6.Visible = True
                        Password.Visible = True
                        SpecPF.BackColor = Color.Gold
                        DeltaPF.BackColor = Color.Red
                        DeltaPF.ForeColor = Color.Gold
                        DeltaPF.Text = "FAILED"
                        TracePF.BackColor = Color.Gold
                    ElseIf GoldenRunPass And Not TraceDeltaPass Then
                        txtprompt.Text = "The Golden Unit Passes, but The trace data has changed. Contact your supervisor"
                        GoldenMode = True
                        btOK.Visible = False
                        btBypass.Visible = True
                        Label6.Visible = True
                        Password.Visible = True
                        SpecPF.BackColor = Color.Gold
                        SpecPF.BackColor = Color.Gold
                        TracePF.BackColor = Color.Red
                        TracePF.ForeColor = Color.Gold
                        TracePF.Text = "FAILED"
                    ElseIf Not GoldenRunPass And DeltaPass And TraceDeltaPass Then
                        txtprompt.Text = "The Golden Unit Fails Specification. Contact your supervisor"
                        txtprompt.ForeColor = Color.Red
                        GoldenMode = True
                        btOK.Visible = False
                        btBypass.Visible = True
                        Label6.Visible = True
                        Password.Visible = True
                        SpecPF.BackColor = Color.Red
                        SpecPF.ForeColor = Color.Gold
                        SpecPF.Text = "FAILED"
                        DeltaPF.BackColor = Color.Gold
                        TracePF.BackColor = Color.Gold
                    ElseIf Not GoldenRunPass And DeltaPass And TraceDeltaPass Then
                        txtprompt.Text = "The Golden Unit Fails Specification. Contact your supervisor"
                        txtprompt.ForeColor = Color.Red
                        GoldenMode = True
                        btOK.Visible = False
                        btBypass.Visible = True
                        Label6.Visible = True
                        Password.Visible = True
                        SpecPF.BackColor = Color.Red
                        SpecPF.ForeColor = Color.Gold
                        SpecPF.Text = "FAILED"
                        DeltaPF.BackColor = Color.Gold
                        TracePF.BackColor = Color.Gold
                    ElseIf Not GoldenRunPass And Not DeltaPass And TraceDeltaPass Then
                        txtprompt.Text = "The Golden Unit Fails Specification and Data delta. Contact your supervisor"
                        txtprompt.ForeColor = Color.Red
                        GoldenMode = True
                        btOK.Visible = False
                        btBypass.Visible = True
                        Label6.Visible = True
                        Password.Visible = True
                        SpecPF.BackColor = Color.Red
                        SpecPF.ForeColor = Color.Gold
                        SpecPF.Text = "FAILED"
                        DeltaPF.BackColor = Color.Red
                        DeltaPF.ForeColor = Color.Gold
                        DeltaPF.Text = "FAILED"
                        TracePF.BackColor = Color.Gold
                    ElseIf Not GoldenRunPass And DeltaPass And Not TraceDeltaPass Then
                        txtprompt.Text = "The Golden Unit Fails Specification and Trace delta. Contact your supervisor"
                        txtprompt.ForeColor = Color.Red
                        GoldenMode = True
                        btOK.Visible = False
                        btBypass.Visible = True
                        Label6.Visible = True
                        Password.Visible = True
                        SpecPF.BackColor = Color.Red
                        SpecPF.ForeColor = Color.Gold
                        SpecPF.Text = "FAILED"
                        DeltaPF.BackColor = Color.Gold
                        TracePF.BackColor = Color.Red
                        TracePF.ForeColor = Color.Gold
                        TracePF.Text = "FAILED"
                    ElseIf Not GoldenRunPass And Not DeltaPass And Not TraceDeltaPass Then
                        txtprompt.Text = "The Golden Unit Fails Everything. Contact your supervisor"
                        txtprompt.ForeColor = Color.Red
                        GoldenMode = True
                        btOK.Visible = False
                        btBypass.Visible = True
                        Label6.Visible = True
                        Password.Visible = True
                        SpecPF.BackColor = Color.Red
                        SpecPF.ForeColor = Color.Gold
                        SpecPF.Text = "FAILED"
                        DeltaPF.BackColor = Color.Red
                        DeltaPF.ForeColor = Color.Gold
                        DeltaPF.Text = "FAILED"
                        TracePF.BackColor = Color.Red
                        TracePF.ForeColor = Color.Gold
                        TracePF.Text = "FAILED"
                    End If
                Else
                    txtprompt.Text = "Test The Golden UUT before you begin"
                    GoldenMode = True
                    btOK.Visible = False
                End If
            Else
                If GoldenRev = "N/A" And txtFixture.Text = "" Then
                    txtprompt.Text = "There is no Test Fixture listed and no Golden UUT data on record. Create both now"
                    btOK.Visible = False
                    btTest.Visible = False
                    btBypass.Visible = True
                    Label6.Visible = True
                    Password.Visible = True
                ElseIf GoldenRev = "N/A" Then
                    txtprompt.Text = "There is a Test Fixture, but no Golden UUT data on record. Create a Golden Unit and test it now"
                    btOK.Visible = False
                    btTest.Visible = False
                    btBypass.Visible = True
                    Label6.Visible = True
                    Password.Visible = True
                Else
                    txtprompt.Text = "There is a Test Fixture and a Golden UUT, but no Golden UUT data on record. Test it now"
                    btOK.Visible = False
                    btTest.Visible = False
                    btBypass.Visible = True
                    Label6.Visible = True
                    Password.Visible = True
                End If
                GoldenMode = True
            End If
        Catch ex As Exception

        End Try
    End Sub
    Private Sub loadpanel()
        If SpecAB_TF Then
            Data4L.Visible = True
            Data4H.Visible = True
            Data4.Visible = False
            Data4L_old.Visible = True
            Data4H_old.Visible = True
            Data4_old.Visible = False
        Else
            Data4L.Visible = False
            Data4H.Visible = False
            Data4.Visible = True
            Data4L_old.Visible = False
            Data4H_old.Visible = False
            Data4_old.Visible = True
        End If
        If ISO_TF Then
            Data3L.Visible = True
            Data3H.Visible = True
            Data3.Visible = False
            Data3L_old.Visible = True
            Data3H_old.Visible = True
            Data3_old.Visible = False
        Else
            Data3L.Visible = False
            Data3H.Visible = False
            Data3.Visible = True
            Data3L_old.Visible = False
            Data3H_old.Visible = False
            Data3_old.Visible = True
        End If
        If IL_TF Then
            Data1L.Visible = True
            Data1H.Visible = True
            Data1.Visible = False
            Data1L_old.Visible = True
            Data1H_old.Visible = True
            Data1_old.Visible = False
        Else
            Data1L_old.Visible = False
            Data1H_old.Visible = False
            Data1_old.Visible = True
            Data1L.Visible = False
            Data1H.Visible = False
            Data1.Visible = True
        End If
    End Sub

    Private Sub TestFixture()
        Dim SQLstr As String
        txtPartNumber.Text = ""
        txtFixture.Text = ""
        txtPlunger.Text = ""
        txtGoldRev.Text = ""
        SQLstr = "SELECT * from TestFixtures where PartNumber = '" & Part & "'"
        If GoldenListDone Then
            txtPartNumber.Text = GoldenPN
            txtFixture.Text = Goldenfixture
            txtPlunger.Text = GoldenPlunger
            txtGoldRev.Text = GoldenRev
            txtFixNum.Text = Goldenfixture_num
        ElseIf CheckforRow(SQLstr, "NetworkSpecs") > 1 Then
            Dim OP As New GoldenList
            OP.StartPosition = FormStartPosition.Manual
            OP.Location = New Point(globals.XLocation + 450, globals.YLocation + 200)
            OP.ShowDialog()
            txtPartNumber.Text = GoldenPN
            txtFixture.Text = Goldenfixture
            txtPlunger.Text = GoldenPlunger
            txtGoldRev.Text = GoldenRev
            txtFixNum.Text = Goldenfixture_num
            GoldenListDone = True
        Else
            If SQLAccess Then
                Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
                Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
                ats.Open()
                System.Threading.Thread.Sleep(10)
                Dim dr As SqlDataReader = cmd.ExecuteReader()
                While Not dr.Read = Nothing
                    txtPartNumber.Text = dr.Item(1)
                    txtFixture.Text = dr.Item(2)
                    txtPlunger.Text = dr.Item(3)
                    txtGoldRev.Text = dr.Item(4)
                    txtFixNum.Text = dr.Item(5)
                    GoldenRev = txtGoldRev.Text
                    ' Compose the picture's file name.
                End While
            End If
        End If
        If txtPartNumber.Text = "" Then
            txtPartNumber.Text = Part
            FixtureSaved = False
        End If
        Dim file_name As String = "\\ippdc\Data\Test Data\Test Department\FINAL TFS\" & txtPartNumber.Text & "_" & txtFixture.Text & ".JPG"
        If Not FileExists(file_name) Then
            file_name = "\\ippdc\Data\Test Data\Test Department\FINAL TFS\no_pic.JPG"
        End If
        ' Load the picture into a Bitmap.
        Dim bm As New Bitmap(file_name)

        ' Display the results.
        picImage.Image = bm
        picImage.SizeMode = PictureBoxSizeMode.StretchImage
    End Sub
    Private Function SaveTestFixture() As Boolean
        Dim SQLstr As String
        If Not txtPartNumber.Text = "" And Not txtFixture.Text = "" And Not txtPlunger.Text = "" And Not txtGoldRev.Text = "" Then
            SQLstr = "SELECT * from TestFixtures where PartNumber = '" & Part & "'"
            If SQL.CheckforRow(SQLstr, "NetworkData") = 0 Then
                SQLstr = "Insert Into Effeciency (FixtureNumber, PartNumber, Plunger, Revision, FixNum) values ('" & txtFixture.Text & "','" & txtPartNumber.Text & "','" & txtPlunger.Text & "','" & txtGoldRev.Text & "'," & txtFixNum.Text & ")"
                SQL.ExecuteSQLCommand(SQLstr, "NetworkSpecs")
            End If

        Else
            MYMsgBox("Please fill in all of the fields", MsgBoxStyle.Critical, "All Data Required!")
            SaveTestFixture = False
            Exit Function
        End If



        If SQLAccess Then
            Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
            Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
            ats.Open()
            System.Threading.Thread.Sleep(10)
            Dim dr As SqlDataReader = cmd.ExecuteReader()
            While Not dr.Read = Nothing
                txtPartNumber.Text = dr.Item(1)
                txtFixture.Text = dr.Item(2)
                txtPlunger.Text = dr.Item(3)
                txtGoldRev.Text = dr.Item(4)
                GoldenRev = txtGoldRev.Text
            End While
        End If
    End Function

    Private Sub TestData()

        Dim SQLStr As String

        SQLStr = "SELECT * from TestData where JobNumber = 'Golden Part' and PartNumber = '" & Part & "' And SerialNumber = '" & GoldenRev & "'"
        If SQL.CheckforRow(SQLStr, "NetworkData") = 0 Then
            GoldenHistory = False
        Else
            GoldenHistory = True
            Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
            Dim cmd As SqlCommand = New SqlCommand(SQLStr, ats)
            ats.Open()
            System.Threading.Thread.Sleep(10)
            Dim dr As SqlDataReader = cmd.ExecuteReader()
            While Not dr.Read = Nothing
                If Not IsDBNull(dr.Item(32)) Then GoldenDate = CStr(dr.Item(32))
                If IL_TF Then
                    If Not IsDBNull(dr.Item(6)) Then ThisData1L_old = dr.Item(6) 'ILH
                    If Not IsDBNull(dr.Item(24)) Then ThisData1H_old = dr.Item(24) 'ILL
                Else
                    If Not IsDBNull(dr.Item(6)) Then ThisData1_old = dr.Item(6) 'IL
                End If
                If Not IsDBNull(dr.Item(7)) Then ThisData2_old = CDbl(dr.Item(7)) 'RL
                If SpecType.Contains("90 DEGREE COUPLER") Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then
                    If ISO_TF Then
                        If Not IsDBNull(dr.Item(9)) Then ThisData3L_old = CStr(dr.Item(9)) 'IsoL
                        If Not IsDBNull(dr.Item(18)) Then ThisData3H_old = CStr(dr.Item(18)) 'IsoH
                    Else
                        If Not IsDBNull(dr.Item(9)) Then ThisData3_old = CStr(dr.Item(9)) 'Iso
                    End If
                    If SpecAB_TF Then
                        If Not IsDBNull(dr.Item(16)) Then ThisData4L_old = CStr(dr.Item(16)) 'AB
                        If Not IsDBNull(dr.Item(17)) Then ThisData4H_old = CStr(dr.Item(17)) 'AB
                    Else
                        If Not IsDBNull(dr.Item(11)) Then ThisData4_old = CStr(dr.Item(11)) 'AB
                    End If

                    If Not IsDBNull(dr.Item(13)) Then ThisData5_old = CStr(dr.Item(13)) 'PB
                Else
                    If Not IsDBNull(dr.Item(8)) Then ThisData3_old = CStr(dr.Item(8)) 'coup
                    If Not IsDBNull(dr.Item(10)) Then ThisData4_old = CStr(dr.Item(10)) 'dir
                    If Not IsDBNull(dr.Item(12)) Then ThisData5_old = CStr(dr.Item(12)) 'cf
                End If

                If Not IsDBNull(dr.Item(6)) Then ThisData1_old = dr.Item(6) 'IL
                If Not IsDBNull(dr.Item(7)) Then ThisData2_old = dr.Item(7) 'RL
                If SpecType.Contains("90 DEGREE COUPLER") Or SpecType.Contains("BALUN") Or SpecType.Contains("COMBINER/DIVIDER") Then
                    If Not IsDBNull(dr.Item(9)) Then ThisData3_old = CStr(dr.Item(9)) 'Iso
                    If Not IsDBNull(dr.Item(11)) Then ThisData4_old = CStr(dr.Item(11)) 'AB
                    If Not IsDBNull(dr.Item(13)) Then ThisData5_old = CStr(dr.Item(13)) 'PB
                Else
                    If Not IsDBNull(dr.Item(9)) Then ThisData3_old = CStr(dr.Item(8)) 'coup
                    If Not IsDBNull(dr.Item(10)) Then ThisData4_old = CStr(dr.Item(10)) 'dir
                    If Not IsDBNull(dr.Item(12)) Then ThisData5_old = CStr(dr.Item(12)) 'cf
                End If

            End While
        End If
    End Sub
    Private Sub Test1()
        Dim DeltaPass As Boolean = True
        PF1.Text = 0
        Test1_limit = 0.06
        If IL_TF Then
            Me.Spec1Max.Text = Format(GetSpecification("InsertionLoss"), "0.00") & "/" & Format(GetSpecification("IL_ex"), "0.00")
            Me.Spec1Min.Text = "N/A"
            If GoldenHistory Then ' There is gold unit data in the database
                If GoldenRunComplete Then ' Recent test completed
                    Dim change1 As Double
                    Dim change2 As Double
                    Dim change1Trace As Double
                    Dim change2Trace As Double
                    If GoldenData1L = "" Or ThisData1L_old = "" Then
                        change1 = 0
                        change2 = 0
                        PF1.Text = 0
                    Else
                        change1 = Math.Abs(CDbl(ThisData1L_old) - CDbl(GoldenData1L))
                        change2 = Math.Abs(CDbl(ThisData1H_old) - CDbl(GoldenData1H))
                        If change1 > change2 Then
                            PF1.Text = TruncateDecimal(change1, 2)
                        Else
                            PF1.Text = TruncateDecimal(change2, 2)
                        End If
                    End If

                    Me.Data1H.Text = GoldenData1H
                    Me.Data1L.Text = GoldenData1L
                    Me.Data1H_old.Text = ThisData1H_old
                    Me.Data1L_old.Text = ThisData1L_old
                    change1Trace = GoldentraceDelta1L
                    change2Trace = GoldentraceDelta1H
                    If change1 > change2 Then
                        Me.PFTrace1.Text = change1
                    Else
                        Me.PFTrace1.Text = change2
                    End If
                Else
                    Me.Data1H.Text = ""
                    Me.Data1L.Text = ""
                    Me.Data1H_old.Text = ThisData1H_old
                    Me.Data1L_old.Text = ThisData1L_old
                End If
            Else
                Me.Data1H.Text = ""
                Me.Data1L.Text = ""
                Me.Data1H_old.Text = ""
                Me.Data1L_old.Text = ""
            End If
        Else
            Me.Spec1Max.Text = Format(GetSpecification("InsertionLoss"), "0.00")
            SpecIL = Me.Spec1Max.Text
            Me.Spec1Min.Text = "N/A"
            If GoldenHistory Then ' There is gold unit data in the database
                If GoldenRunComplete Then ' Recent test completed
                    Dim change As Double
                    If GoldenData1 = "" Or ThisData1_old = "" Then
                        change = 0
                        PF1.Text = 0
                    Else
                        change = Math.Abs(CDbl(ThisData1_old) - CDbl(GoldenData1))
                        Me.Data1.Text = GoldenData1
                        Me.Data1_old.Text = ThisData1_old
                        PF1.Text = TruncateDecimal(change, 2)
                        Me.PFTrace1.Text = GoldentraceDelta1
                    End If
                   
                Else
                    Me.Data1.Text = ""
                    Me.Data1_old.Text = ThisData1_old
                End If
                Else
                    Me.Data1.Text = ""
                    Me.Data1_old.Text = ""
                End If
        End If
    End Sub
    Private Sub Test2()
        PF2.Text = 0
        Me.Spec2Min.Text = "N/A"
        Test2_limit = 1.5
        If Not IsDBNull(GetSpecification("VSWR")) Then
            Me.Spec2Max.Text = Format(GetSpecification("VSWR"), "0.0")
            SpecTest2 = GetSpecification("VSWR")
        End If
        If GoldenHistory Then ' There is gold unit data in the database
            If GoldenRunComplete Then ' Recent test completed
                Dim change As Double
                If GoldenData2 = "" Or ThisData2_old = "" Then
                    change = 0
                    PF2.Text = 0
                Else
                    change = TruncateDecimal(Math.Abs(CDbl(ThisData2_old) - CDbl(GoldenData2)), 2)
                    PF2.Text = change
                End If

                Me.Data2.Text = GoldenData2
                Me.Data2_old.Text = ThisData2_old
                Me.PFTrace2.Text = GoldentraceDelta2
            Else
                Me.Data2.Text = ""
                Me.Data2_old.Text = ThisData2_old
            End If
        Else
            Me.Data2.Text = ""
            Me.Data2_old.Text = ""
        End If
    End Sub
    Private Sub Test3()
        PF3.Text = 0
        If SpecIndex = 0 Or SpecIndex = 1 Or SpecIndex = 3 Then
            If Not IsDBNull(GetSpecification("Isolation")) And ISO_TF Then
                Test3_limit = 1.5
                TestLabel3.Text = "Isolation D/B:  dB"
                SpecISOL = Format(GetSpecification("IsolationL"), "0.0")
                SpecISOH = Format(GetSpecification("IsolationH"), "0.0")
                Me.Spec3Max.Text = 0 - SpecISOL & "/" & 0 - SpecISOH
                If GoldenHistory Then ' There is gold unit data in the database
                    If GoldenRunComplete Then ' Recent test completed
                        Dim change1 As Double
                        Dim change2 As Double
                        Dim change1Trace As Double
                        Dim change2Trace As Double
                        If GoldenData3L = "" Or ThisData3L_old = "" Then
                            change1 = 0
                            change2 = 0
                            PF3.Text = 0
                        Else
                            change1 = Math.Abs(CDbl(ThisData3L_old) - CDbl(GoldenData3L))
                            change2 = Math.Abs(CDbl(ThisData3H_old) - CDbl(GoldenData3H))
                            If change1 > change2 Then
                                PF3.Text = TruncateDecimal(change1, 2)
                            Else
                                PF3.Text = TruncateDecimal(change2, 2)
                            End If
                        End If
                        Me.Data3H.Text = GoldenData3H
                        Me.Data3L.Text = GoldenData3L
                        Me.Data3H_old.Text = ThisData3H_old
                        Me.Data3L_old.Text = ThisData3L_old
                        change1Trace = GoldentraceDelta3L
                        change2Trace = GoldentraceDelta3H
                        If change1 > change2 Then
                            Me.PFTrace3.Text = change1
                        Else
                            Me.PFTrace3.Text = change2
                        End If
                    Else
                        Me.Data3H.Text = ""
                        Me.Data3L.Text = ""
                        Me.Data3H_old.Text = ThisData3H_old
                        Me.Data3L_old.Text = ThisData3L_old
                    End If
                Else
                    Me.Data3H.Text = ""
                    Me.Data3.Text = ""
                    Me.Data3H_old.Text = ""
                    Me.Data3L_old.Text = ""
                End If
            Else
                TestLabel3.Text = "Isolation:  dB"
                Test3_limit = 1.5
                Me.Spec3Max.Text = Format(0 - GetSpecification("Isolation"), "0.0")
                SpecTest3 = Format(GetSpecification("Isolation"), "0.0")
                SpecISO = SpecTest3
                If GoldenHistory Then ' There is gold unit data in the database
                    If GoldenRunComplete Then ' Recent test completed
                        Dim change As Double
                        If GoldenData3 = "" Or ThisData3_old = "" Then
                            change = 0
                            PF3.Text = 0
                        Else
                            change = Math.Abs(CDbl(ThisData3_old) - CDbl(GoldenData3))
                            PF3.Text = change
                        End If


                        Me.Data3.Text = GoldenData3
                        Me.Data3_old.Text = ThisData3_old
                        Me.PFTrace3.Text = GoldentraceDelta3
                    Else
                        Me.Data3.Text = ""
                        Me.Data3_old.Text = ThisData3_old
                    End If
                Else
                    Me.Data3.Text = ""
                    Me.Data3_old.Text = ""
                End If
            End If
            Me.Spec3Min.Text = "N/A"
        ElseIf SpecIndex = 2 Then
            TestLabel3.Text = "Coupling:  dB"
            Test3_limit = 1.0
            If Not IsDBNull(GetSpecification("Coupling")) Then
                Me.Spec3Min.Text = Format(GetSpecification("Coupling") - GetSpecification("CoupPlusMinus"), "0.0")
                Me.Spec3Max.Text = Format(GetSpecification("Coupling") + GetSpecification("CoupPlusMinus"), "0.0")
                SpecCOUP = GetSpecification("Coupling")
            End If
            If GoldenHistory Then ' There is gold unit data in the database
                If GoldenRunComplete Then ' Recent test completed
                    Dim change As Double
                    If GoldenData3 = "" Or ThisData3_old = "" Then
                        change = 0
                        PF3.Text = 0
                    Else
                        change = Math.Abs(CDbl(ThisData3_old) - CDbl(GoldenData3))
                        PF3.Text = TruncateDecimal(change, 2)
                    End If


                    Me.Data3.Text = GoldenData3
                    Me.Data3_old.Text = ThisData3_old
                    Me.PFTrace3.Text = GoldentraceDelta3
                Else
                    Me.Data3.Text = ""
                    Me.Data3_old.Text = ThisData3_old
                End If
            Else
                Me.Data3.Text = ""
                Me.Data3_old.Text = ""
            End If
        End If
    End Sub
    Private Sub Test4()
        PF4.Text = 0
        If SpecIndex = 0 Or SpecIndex = 1 Or SpecIndex = 3 Then
            TestLabel4.Text = "Amplitude Balance dB"
            Test4_limit = 0.6
            If Not IsDBNull(GetSpecification("AmplitudeBalance")) Then
                Me.Spec4Min.Text = Format(GetSpecification("AmplitudeBalance"), "0.00")
                If SpecAB_TF Then
                    TestLabel4.Text = "Amplitude Balance D/B  dB"
                    SpecTest4 = Me.Spec4Min.Text
                    Me.Spec4Min.Text = Me.Spec4Min.Text & "/" & SpecAB_exp
                    Me.Spec4Max.Text = Me.Spec4Min.Text
                    SpecTest4_exp = SpecAB_exp
                    If GoldenHistory Then ' There is gold unit data in the database
                        If GoldenRunComplete Then ' Recent test completed
                            Dim change1 As Double
                            Dim change2 As Double
                            Dim change1Trace As Double
                            Dim change2Trace As Double
                            If GoldenData4L = "" Or ThisData4L_old = "" Then
                                change1 = 0
                                change2 = 0
                                PF4.Text = 0
                            Else
                                change1 = Math.Abs(CDbl(ThisData4L_old) - CDbl(GoldenData4L))
                                change2 = Math.Abs(CDbl(ThisData4H_old) - CDbl(GoldenData4H))
                                If change1 > change2 Then
                                    PF4.Text = TruncateDecimal(change1, 2)
                                Else
                                    PF4.Text = TruncateDecimal(change2, 2)
                                End If
                            End If

                            Me.Data4H.Text = GoldenData4H
                            Me.Data4L.Text = GoldenData4L
                            Me.Data4H_old.Text = ThisData4H_old
                            Me.Data4L_old.Text = ThisData4L_old
                            change1Trace = GoldentraceDelta4L
                            change2Trace = GoldentraceDelta4H
                            If change1 > change2 Then
                                Me.PFTrace4.Text = change1
                            Else
                                Me.PFTrace4.Text = change2
                            End If
                        Else
                            Me.Data4H.Text = ""
                            Me.Data4L.Text = ""
                            Me.Data4H_old.Text = ThisData4H_old
                            Me.Data4L_old.Text = ThisData4L_old
                        End If
                        Else
                            Me.Data4H.Text = ""
                            Me.Data4.Text = ""
                            Me.Data4H_old.Text = ""
                            Me.Data4L_old.Text = ""
                        End If
                Else
                    SpecTest4 = Me.Spec4Min.Text
                    Me.Spec4Max.Text = Me.Spec4Min.Text
                    If GoldenHistory Then ' There is gold unit data in the database
                        If GoldenRunComplete Then ' Recent test completed
                            Dim change As Double
                            If GoldenData4 = "" Or ThisData4_old = "" Then
                                change = 0
                                PF4.Text = 0
                            Else
                                change = Math.Abs(CDbl(ThisData4_old) - CDbl(GoldenData4))
                                PF4.Text = TruncateDecimal(change, 2)
                            End If

                            Me.Data4.Text = GoldenData4
                            Me.Data4_old.Text = ThisData4_old
                            Me.PFTrace4.Text = GoldentraceDelta4
                        Else
                            Me.Data4.Text = ""
                            Me.Data4_old.Text = ThisData4_old
                        End If
                        Else
                            Me.Data4.Text = ""
                            Me.Data4_old.Text = ""
                        End If
                End If
            End If
        ElseIf SpecIndex = 2 Then
            TestLabel4.Text = "Directivity: dB"
            Test4_limit = 1.5
            If Not IsDBNull(GetSpecification("Directivity")) Then
                Me.Spec4Min.Text = Format(GetSpecification("Directivity"), "0.0")
                SpecDIRECT = Me.Spec4Min.Text
            End If
            Me.Spec4Max.Text = "N/A"
            If GoldenHistory Then ' There is gold unit data in the database
                If GoldenRunComplete Then ' Recent test completed
                    Dim change As Double
                    If GoldenData4 = "" Or ThisData4_old = "" Then
                        change = 0
                        PF4.Text = 0
                    Else
                        change = Math.Abs(CDbl(ThisData4_old) - CDbl(GoldenData4))
                        PF4.Text = TruncateDecimal(change, 2)
                    End If
                    
                    Me.Data4.Text = GoldenData4
                    Me.Data4_old.Text = ThisData4_old
                    Me.PFTrace4.Text = GoldentraceDelta1
                Else
                    Me.Data4.Text = ""
                    Me.Data4_old.Text = ThisData4_old
                End If
            Else
                Me.Data4.Text = ""
                Me.Data4_old.Text = ""
            End If
        End If

    End Sub
    Private Sub Test5()
        If SpecIndex = 0 Or SpecIndex = 1 Or SpecIndex = 3 Then
            TestLabel5.Text = "Phase Balance: Deg"
            Test5_limit = 1.5
            Dim Temp As String = Format(GetSpecification("PhaseBalance"), "0.0")

            If Not IsDBNull(GetSpecification("PhaseBalance")) Then Me.Spec5Min.Text = Format(GetSpecification("PhaseBalance"), "0.0")
            Me.Data5.Text = ""
            If Not IsDBNull(GetSpecification("PhaseBalance")) Then
                Me.Spec5Max.Text = Format(GetSpecification("PhaseBalance"), "0.0")
                Me.Spec5Min.Text = Me.Spec5Max.Text
                SpecTest5 = Me.Spec5Min.Text
            End If
            Dim change As Double
            If GoldenHistory Then ' There is gold unit data in the database
               If GoldenRunComplete Then ' Recent test completed
                    If GoldenData5 = "" Or ThisData5_old = "" Then
                        PF5.Text = 0
                    Else
                        change = Math.Abs(CDbl(ThisData5_old) - CDbl(GoldenData5))
                        PF5.Text = change
                    End If

                    Me.Data5.Text = GoldenData5
                    Me.Data5_old.Text = ThisData5_old
                    Me.PFTrace5.Text = GoldentraceDelta5
                Else
                    Me.Data5.Text = ""
                    Me.Data5_old.Text = ThisData5_old
                End If
            End If
        ElseIf SpecIndex = 2 Then
            TestLabel5.Text = "Coupled Flatness dB"
            Test5_limit = 0.5
            Me.Data5.Text = ""
            Dim change As Double
            If Not IsDBNull(GetSpecification("CoupledFlatness")) Then
                Me.Spec5Max.Text = Format(GetSpecification("CoupledFlatness"), "0.00")
                Me.Spec5Min.ForeColor = Color.CornflowerBlue
                Me.Spec5Min.Text = Format(GetSpecification("CoupledFlatness"), "0.00")
                Me.Spec5Max.ForeColor = Color.CornflowerBlue
                SpecCOUPFLAT = Me.Spec5Min.Text
            End If
            If GoldenHistory Then ' There is gold unit data in the database
                If GoldenRunComplete Then ' Recent test completed
                    If GoldenData5 = "" Or ThisData5_old = "" Then
                        PF5.Text = 0
                    Else
                        change = Math.Abs(CDbl(ThisData5_old) - CDbl(GoldenData5))
                        PF5.Text = change
                    End If

                    Me.Data5.Text = GoldenData5
                    Me.Data5_old.Text = ThisData5_old
                    Me.PFTrace5.Text = GoldentraceDelta5
                Else
                    Me.Data5.Text = ""
                    Me.Data5_old.Text = ThisData5_old
                End If
            Else
                Me.Data5.Text = ""
                Me.Data5_old.Text = ""
            End If
        End If
    End Sub
    Private Sub resetdata()
        GoldenData1 = ""
        GoldenData1H = ""
        GoldenData1L = ""
        GoldenData2 = ""
        GoldenData3 = ""
        GoldenData3H = ""
        GoldenData3L = ""
        GoldenData4 = ""
        GoldenData4L = ""
        GoldenData4H = ""
        GoldenData5 = ""
        GoldenData5H = ""
        GoldenData5L = ""
    End Sub
    Private Sub LoadSpecs()
        Dim SQLstr As String
        SQLstr = "SELECT * from Specifications where PartNumber = '" & Part & "'"
        If SQLAccess Then
            Dim ats As SqlConnection = New SqlConnection(SQLConnStr)
            Dim cmd As SqlCommand = New SqlCommand(SQLstr, ats)
            ats.Open()
            System.Threading.Thread.Sleep(10)
            Dim dr As SqlDataReader = cmd.ExecuteReader()
            While Not dr.Read = Nothing
                If Not IsDBNull(dr.Item(1)) Then
                    If dr.Item(1) = "90 DEGREE COUPLER" Or dr.Item(1) = "90 DEGREE COUPLER SMD" Then
                        If dr.Item(1) = "90 DEGREE COUPLER SMD" Then
                            SMD = True
                        Else
                            SMD = False
                        End If
                        SpecIndex = 0
                        SpecType = "90 DEGREE COUPLER"
                        TestLabel3.Visible = True
                        Spec3Min.Visible = True
                        Spec3Max.Visible = True
                        Data3.Visible = True
                        PF3.Visible = True
                    ElseIf dr.Item(1) = "TRANSFORMER" Or dr.Item(1) = "TRANSFORMER SMD" Then
                        If dr.Item(1) = "TRANSFORMER SMD" Then
                            SMD = True
                        Else
                            SMD = False
                        End If
                        IL_TF = dr.Item(91)
                        SpecIndex = 1
                        SpecType = "TRANSFORMER"

                        TestLabel3.Visible = False
                        Spec3Min.Visible = False
                        Spec3Max.Visible = False
                        Data3.Visible = False
                        PF3.Visible = False
                        TestLabel4.Visible = False
                        Spec4Min.Visible = False
                        Spec4Max.Visible = False
                        Data4.Visible = False
                        PF4.Visible = False
                        TestLabel5.Visible = False
                        Spec5Min.Visible = False
                        Spec5Max.Visible = False
                        Data5.Visible = False
                        PF5.Visible = False
                    ElseIf dr.Item(1) = "BALUN" Or dr.Item(1) = "BALUN SMD" Then
                        If dr.Item(1) = "BALUN SMD" Then
                            SMD = True
                        Else
                            SMD = False
                        End If
                        SpecIndex = 1
                        SpecType = "BALUN"
                        TestLabel3.Visible = False
                        Spec3Min.Visible = False
                        Spec3Max.Visible = False
                        Data3.Visible = False
                        PF3.Visible = False
                    ElseIf dr.Item(1) = "SINGLE DIRECTIONAL COUPLER" Or dr.Item(1) = "SINGLE DIRECTIONAL COUPLER SMD" Then
                        If dr.Item(1) = "SINGLE DIRECTIONAL COUPLER SMD" Then
                            SMD = True
                        Else
                            SMD = False
                        End If
                        SpecIndex = 2
                        SpecType = "SINGLE DIRECTIONAL COUPLER"
                        TestLabel3.Visible = True
                        Spec3Min.Visible = True
                        Spec3Max.Visible = True
                        Data3.Visible = True
                        PF3.Visible = True
                    ElseIf dr.Item(1) = "DUAL DIRECTIONAL COUPLER" Or dr.Item(1) = "DUAL DIRECTIONAL COUPLER SMD" Then
                        If dr.Item(1) = "DUAL DIRECTIONAL COUPLER SMD" Then
                            SMD = True
                        Else
                            SMD = False
                        End If
                        SpecIndex = 2
                        SwitchPorts = dr.Item(100)
                        SpecType = "DUAL DIRECTIONAL COUPLER"
                        TestLabel3.Visible = True
                        Spec3Min.Visible = True
                        Spec3Max.Visible = True
                        Data3.Visible = True
                        PF3.Visible = True
                    ElseIf dr.Item(1) = "BI DIRECTIONAL COUPLER" Or dr.Item(1) = "BI DIRECTIONAL COUPLER SMD" Then
                        If dr.Item(1) = "BI DIRECTIONAL COUPLER SMD" Then
                            SMD = True
                        Else
                            SMD = False
                        End If
                        SpecIndex = 2
                        SpecType = "BI DIRECTIONAL COUPLER"
                        TestLabel3.Visible = True
                        Spec3Min.Visible = True
                        Spec3Max.Visible = True
                        Data3.Visible = True
                        PF3.Visible = True
                    ElseIf dr.Item(1) = "COMBINER/DIVIDER" Or dr.Item(1) = "COMBINER/DIVIDER SMD" Then
                        If dr.Item(1) = "COMBINER/DIVIDER SMD" Then
                            SMD = True
                        Else
                            SMD = False
                        End If
                        SpecIndex = 3
                        TestLabel3.Visible = True
                        Spec3Min.Visible = True
                        Spec3Max.Visible = True
                        Data3.Visible = True
                        PF3.Visible = True
                        SpecType.Contains("COMBINER/DIVIDER")
                        'If dr.Item(13) = 0 Then ckTest3.Checked = False
                    End If
                Else
                    SpecIndex = 0
                    SpecType = "90 DEGREE COUPLER"
                    TestLabel3.Visible = True
                    Spec3Min.Visible = True
                    Spec3Max.Visible = True
                    Data3.Visible = True
                    PF3.Visible = True
                    SMD = False
                End If
            End While
            ats.Close()
        End If
    End Sub

    Private Sub btBypass_Click(sender As Object, e As EventArgs) Handles btBypass.Click
        If Password.Text = "Testlog2020" Then
            GoldenMode = False
            GoldenModeBypass = True
            Me.Close()
        Else
            MYMsgBox("Enter the Supervisor Password", MsgBoxStyle.Critical, "Password Required")
        End If

    End Sub

    Private Sub btGoldenData_Click(sender As Object, e As EventArgs) Handles btGoldenData.Click
        If Not FixtureSaved Then
            SaveTestFixture()
            If SaveTestFixture() = False Then
                Exit Sub
            End If
        End If
        GoldenMode = True
        GoldenData = False
        resetdata()
        Me.Close()
    End Sub
    Private Sub btOK_Click(sender As Object, e As EventArgs) Handles btOK.Click
        GoldenModeBypass = True
        GoldenMode = False
        GoldenData = False
        Me.Close()
    End Sub

    Private Sub btTest_Click(sender As Object, e As EventArgs) Handles btTest.Click
        GoldenModeBypass = False
        GoldenMode = True
        GoldenData = True
        resetdata()
        Me.Close()
    End Sub

    
    Private Sub picImage_Click(sender As Object, e As EventArgs) Handles picImage.Click
        Dim url As String
        url = "http://inn-sqlexpress:8888/testfixtures/" & "?part_num=" & txtPartNumber.Text & "&fix_num=" & txtFixture.Text

        Process.Start("msedge.exe", url)
    End Sub

   
End Class