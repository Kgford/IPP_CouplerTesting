Imports System.Data
Imports System.Data.SqlClient
Imports System.Reflection



Public Class MYMsg
    Public Sub New()
        InitializeComponent()
        Try
            txttitle.Text = MYMSG_TITLE
            txtprompt.Text = MYMSG_PROMPT
            If MYMSG_BTN = 0 Then 'OK Only
                btOK.Location = New Point(165, 65) 'Center
                btOK.Visible = True

                btYES.Visible = False
                btNO.Visible = False
                btCANCEL.Visible = False
                btABORT.Visible = False
                btRETRY.Visible = False
                btIGNORE.Visible = False
            ElseIf MYMSG_BTN = 1 Then 'Display OK and Cancel buttons
                btOK.Location = New Point(96, 65) 'two buttons
                btCANCEL.Location = New Point(220, 65) 'two buttons
                btOK.Visible = True
                btCANCEL.Visible = True

                btYES.Visible = False
                btNO.Visible = False
                btABORT.Visible = False
                btRETRY.Visible = False
                btIGNORE.Visible = False
            ElseIf MYMSG_BTN = 2 Then ''Display Abort, Retry, and Ignore buttons
                btABORT.Location = New Point(39, 65) 'Left
                btRETRY.Location = New Point(165, 65) 'Center
                btIGNORE.Location = New Point(291, 65) 'Right
                btABORT.Visible = True
                btRETRY.Visible = True
                btIGNORE.Visible = True

                btYES.Visible = False
                btOK.Visible = False
                btNO.Visible = False
                btCANCEL.Visible = False
            ElseIf MYMSG_BTN = 3 Then ''Display Yes, No, and Cancel buttons
                btYES.Location = New Point(39, 65) 'Left
                btNO.Location = New Point(165, 65) 'Center
                btCANCEL.Location = New Point(291, 65) 'Right
                btYES.Visible = True
                btNO.Visible = True
                btCANCEL.Visible = True

                btABORT.Visible = False
                btOK.Visible = False
                btRETRY.Visible = False
                btIGNORE.Visible = False
            ElseIf MYMSG_BTN = 4 Then 'Display Yes and No buttons
                btYES.Location = New Point(96, 65) 'two buttons
                btNO.Location = New Point(220, 65) 'two buttons
                btYES.Visible = True
                btNO.Visible = True

                btOK.Visible = False
                btCANCEL.Visible = False
                btABORT.Visible = False
                btRETRY.Visible = False
                btIGNORE.Visible = False
            ElseIf MYMSG_BTN = 5 Then 'Display Retry and Cancel buttons
                btRETRY.Location = New Point(96, 65) 'two buttons
                btCANCEL.Location = New Point(220, 65) 'two buttons
                btRETRY.Visible = True
                btCANCEL.Visible = True

                btOK.Visible = False
                btNO.Visible = False
                btABORT.Visible = False
                btYES.Visible = False
                btIGNORE.Visible = False
            End If

        Catch

        End Try
    End Sub

    Private Sub btrev_Click(sender As Object, e As EventArgs) Handles btOK.Click
        MYMSG_RTN = vbOK
        Me.Close()

    End Sub

    Private Sub btNO_Click(sender As Object, e As EventArgs) Handles btNO.Click
        MYMSG_RTN = vbNo
        Me.Close()
    End Sub

    Private Sub btRETRY_Click(sender As Object, e As EventArgs) Handles btRETRY.Click
        MYMSG_RTN = vbRetry
        Me.Close()
    End Sub

    Private Sub btYES_Click(sender As Object, e As EventArgs) Handles btYES.Click
        MYMSG_RTN = vbYes
        Me.Close()
    End Sub

    Private Sub btABORT_Click(sender As Object, e As EventArgs) Handles btABORT.Click
        MYMSG_RTN = vbAbort
        Me.Close()
    End Sub

    Private Sub btCANCEL_Click(sender As Object, e As EventArgs) Handles btCANCEL.Click
        MYMSG_RTN = vbCancel
        Me.Close()
    End Sub

    Private Sub btIGNORE_Click(sender As Object, e As EventArgs) Handles btIGNORE.Click
        MYMSG_RTN = vbIgnore
        Me.Close()
    End Sub
End Class