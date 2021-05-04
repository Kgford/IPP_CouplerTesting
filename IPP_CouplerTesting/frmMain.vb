Public Class frmMain
    'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
    'Const EM_UNDO = &HC7
    'Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
    'Dim frmD As frmDocument
    'Dim frmB As New frmBrowser
    Dim BrowserOn As Boolean
    Private Sub MDIForm_Load()

        Me.Left = -60
        Me.Top = -60
        Me.Width = 25320
        Me.Height = 15420


        'Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
        'Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
        'Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
        'Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
        'txtVersion.Text = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        'LoadNewDoc()
        'Me.Arrange vbTileHorizontal
        'mnuAutotest_Click()
    End Sub


    Private Sub LoadNewDoc()
        'Static lDocumentCount As Long
        '    lDocumentCount = lDocumentCount + 1
        '    Set frmD = New frmDocument
        '    frmD.Caption = "IPP Notes: " & lDocumentCount
        '    frmD.Show

    End Sub


    Private Sub MDIForm_Unload(Cancel As Integer)
        'If Me.WindowState <> vbMinimized Then
        '    SaveSetting(App.Title, "Settings", "MainLeft", Me.Left)
        '    SaveSetting(App.Title, "Settings", "MainTop", Me.Top)
        '    SaveSetting(App.Title, "Settings", "MainWidth", Me.Width)
        '    SaveSetting(App.Title, "Settings", "MainHeight", Me.Height)
        'End If
    End Sub

    Private Sub Form_Unload(Cancel As Integer)

        End
    End Sub

   


    Private Sub ScanGPIBToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ScanGPIBToolStripMenuItem.Click

    End Sub

    Private Sub SpecificationsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpecificationsToolStripMenuItem.Click
        'frmSpecifications.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
        'frmSpecifications.Top = GetSetting(App.Title, "Settings", "MainTop", 1000) + 2800
        'frmSpecifications.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
        'frmSpecifications.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500) - 2000
        'frmSpecifications.Show(vbModeless, Me)
    End Sub

    Private Sub ScreenShotToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ScreenShotToolStripMenuItem.Click

    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click

    End Sub

    Private Sub AutotestToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles AutotestToolStripMenuItem2.Click

    End Sub
End Class