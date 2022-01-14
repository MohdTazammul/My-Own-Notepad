Imports System
Imports System.IO
Imports System.IO.Path
Public Class Form1
    Dim filename As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.RichTextBox1.Select()
        If filename = "" Then
            Me.Text = "Tazammul's Notepad ~ Untitled"
        Else
            Me.Text = "Tazammul's Notepad ~ " & filename
        End If
        ToolStripTextBox2.Text = (Me.Opacity) * 100
    End Sub

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        RichTextBox1.Text = ""
        RichTextBox1.ZoomFactor = RichTextBox1.ZoomFactor + 1
    End Sub

    Private Sub NewWindowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewWindowToolStripMenuItem.Click
        Dim form As New Form1
        form.Show()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim oReader As New StreamReader(OpenFileDialog1.FileName, True)
        RichTextBox1.Text = oReader.ReadToEnd()
        filename = GetFileName(OpenFileDialog1.FileName)
        Me.Text = "Tazammul's Notepad ~ " & filename
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        SaveFileDialog1.ShowDialog()
    End Sub

    Private Sub SaveFileDialog1_FileOk(sender As Object, e As ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        My.Computer.FileSystem.WriteAllText(SaveFileDialog1.FileName, RichTextBox1.Text, False)
        filename = GetFileName(SaveFileDialog1.FileName)
        Me.Text = "Tazammul's Notepad ~ " & filename
    End Sub

    Private Sub PageSetupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PageSetupToolStripMenuItem.Click
        PageSetupDialog1.Document = PrintDocument1
        PageSetupDialog1.Document.DefaultPageSettings.Color = False
        PageSetupDialog1.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        e.Graphics.DrawString(RichTextBox1.Text, RichTextBox1.Font, Brushes.Black, 100, 100)
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        PrintDialog1.Document = PrintDocument1
        If PrintDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If RichTextBox1.Text = "" Then
            End
        Else
            Dim op As Integer = MsgBox("Do you want to save changes to " & filename & "?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.DefaultButton1 + MsgBoxStyle.Exclamation + MsgBoxStyle.ApplicationModal)
            If (op = 6) Then
                SaveToolStripMenuItem.PerformClick()
            ElseIf (op = 7) Then
                End
            End If
        End If
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        RichTextBox1.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click
        RichTextBox1.Redo()
    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        RichTextBox1.SelectedText = ""
    End Sub

    Private Sub FindToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindToolStripMenuItem.Click
        Dim str As String = InputBox("What to find", "Find", "Enter String")

        Dim startText As Integer = 0
        Dim endText As Integer

        endText = RichTextBox1.Text.LastIndexOf(str)

        ' RichTextBox1.SelectAll()
        'RichTextBox1.SelectionBackColor = Color.White


        While startText < endText

            RichTextBox1.Find(str, startText, RichTextBox1.TextLength, RichTextBoxFinds.MatchCase)
            RichTextBox1.SelectionBackColor = Color.Brown

            startText = RichTextBox1.Text.IndexOf(str, startText) + 1


        End While




    End Sub

    Private Sub ReplaceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReplaceToolStripMenuItem.Click

    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        RichTextBox1.SelectAll()
    End Sub

    Private Sub DateTimeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DateTimeToolStripMenuItem.Click
        Dim dt As Date = Today
        RichTextBox1.AppendText(Now.ToString("hh:mm:ss tt, dddd, M MMM, yyyy"))
    End Sub

    Private Sub FontStyleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontStyleToolStripMenuItem.Click
        If FontDialog1.ShowDialog = DialogResult.OK Then
            RichTextBox1.SelectionFont = FontDialog1.Font
        End If
    End Sub

    Private Sub FontDialog1_Apply(sender As Object, e As EventArgs) Handles FontDialog1.Apply
        RichTextBox1.SelectionFont = FontDialog1.Font
    End Sub

    Private Sub FontColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontColorToolStripMenuItem.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.SelectionColor = ColorDialog1.Color
    End Sub

    Private Sub BacgroundColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BacgroundColorToolStripMenuItem.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.BackColor = ColorDialog1.Color
    End Sub

    Private Sub ToolStripTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Me.Opacity = Val(ToolStripTextBox2.Text) / 100
        End If
    End Sub

    Private Sub WordWrapToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WordWrapToolStripMenuItem.Click
        If RichTextBox1.WordWrap = True Then
            RichTextBox1.WordWrap = False
        Else
            RichTextBox1.WordWrap = True
        End If
    End Sub

    Private Sub BulletToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BulletToolStripMenuItem.Click
        If RichTextBox1.SelectionBullet = True Then
            RichTextBox1.SelectionBullet = False
        Else
            RichTextBox1.SelectionBullet = True
        End If
    End Sub

    Private Sub ZoomFactorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZoomFactorToolStripMenuItem.Click
        RichTextBox1.ZoomFactor = 2
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        RichTextBox1.ZoomFactor = 3
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        RichTextBox1.ZoomFactor = 4
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        RichTextBox1.ZoomFactor = 5
    End Sub

    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        RichTextBox1.ZoomFactor = 6
    End Sub

    Private Sub ViewHelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewHelpToolStripMenuItem.Click
        MsgBox("Koi Help Welp ni he, khud dekh lo kese use krna he",, "Aaye bade")
    End Sub

    Private Sub AboutTextEditorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutTextEditorToolStripMenuItem.Click
        MsgBox("This software is developed by Mohd Tazammul, please if you have any feedback regarding the software positive or negative don't hesitate just say to me. Thanks for using my Notepad", MsgBoxStyle.Information, "My Own Notepad")
    End Sub

    Private Sub LeftToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LeftToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Left
    End Sub

    Private Sub CenterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CenterToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Center
    End Sub

    Private Sub RightToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RightToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Right
    End Sub

    Private Sub ToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem7.Click
        If (RichTextBox1.SelectionBackColor <> RichTextBox1.BackColor Or RichTextBox1.SelectionColor <> Color.White) Then
            RichTextBox1.SelectionBackColor = RichTextBox1.BackColor
            RichTextBox1.ForeColor = Color.White
        Else
            RichTextBox1.SelectionBackColor = Color.Black
            RichTextBox1.ForeColor = Color.White
        End If
    End Sub

    Private Sub HighlightColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HighlightColorToolStripMenuItem.Click
        ColorDialog1.ShowDialog()
        RichTextBox1.SelectionBackColor = ColorDialog1.Color
        RichTextBox1.SelectionColor = Color.White
    End Sub

    Private Sub DarkModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DarkModeToolStripMenuItem.Click
        Dim bc, fc As Color
        If (RichTextBox1.BackColor <> Color.Black And RichTextBox1.ForeColor <> Color.White) Then
            bc = RichTextBox1.BackColor
            fc = RichTextBox1.ForeColor
        End If
        If (RichTextBox1.BackColor = Color.Black And RichTextBox1.ForeColor = Color.White) Then
            RichTextBox1.BackColor = bc
            RichTextBox1.ForeColor = fc
        Else
            RichTextBox1.BackColor = Color.Black
            RichTextBox1.ForeColor = Color.White
        End If
    End Sub

    Private Sub TransparentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransparentToolStripMenuItem.Click
        Me.TransparencyKey = RichTextBox1.BackColor
    End Sub
End Class
