
Imports System.IO
Public Class Form3
    Private ReadOnly saveFile As New Windows.Forms.SaveFileDialog
    Private ReadOnly PrintFile As New Windows.Forms.PrintDialog

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.MaximizeBox = True
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.Sizable
        Me.StartPosition = FormStartPosition.CenterScreen
        RichTextBox1.BackColor = Color.GhostWhite
        RichTextBox1.ForeColor = Color.Black
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Try

            OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop


            OpenFileDialog1.Filter = "Text Documents|*.txt|All Files|*.*"

            OpenFileDialog1.Title = "Se deschide:"
            If (OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK) Then
                RichTextBox1.Text = File.ReadAllText(OpenFileDialog1.FileName)
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Try

            If RichTextBox1.Text.Length > 0 And RichTextBox1.Text <> "" Then

                saveFile.Filter = "Text Documents|*.txt|All Files|*.*"
                saveFile.Title = "Se salveaza la:"
                saveFile.ShowDialog()
                File.WriteAllText(saveFile.FileName, RichTextBox1.Text)
                Me.Text = saveFile.FileName.ToString & " - " & "Proiect"
            Else
                MsgBox("Eroare: Nu exista text pentru a putea fi salvat" & Environment.NewLine & "")
            End If
        Catch ex As Exception
            MsgBox("" & Environment.NewLine & ex.Message())
        End Try
    End Sub
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        End
    End Sub

    Private Sub CutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem.Click
        If RichTextBox1.Text.Length > 0 Then
            RichTextBox1.Cut()
        Else
            MsgBox("Va rugam sa selectati textul")
        End If

    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click

        Try
            RichTextBox1.Paste()
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

    End Sub

    Private ReadOnly fontStyle As New Windows.Forms.FontDialog
    Private Sub FontToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontToolStripMenuItem.Click
        fontStyle.ShowDialog()
        RichTextBox1.Font = fontStyle.Font
    End Sub

    Private ReadOnly addColor As New Windows.Forms.ColorDialog
    Private Sub ColorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ColorToolStripMenuItem.Click
        addColor.ShowDialog()
        RichTextBox1.ForeColor = addColor.Color
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        Try
            RichTextBox1.Copy()
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try

    End Sub

    Private Sub SelectToolStripMenuItem_Click(sender As Object, e As EventArgs)
        If RichTextBox1.Text.Length > 0 Then
            RichTextBox1.SelectAll()
        Else
            MsgBox("Nu este nimic selectat!")
        End If
    End Sub


    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        Dim newFile As New Form3
        newFile.ShowDialog()
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs)
        If (OpenFileDialog1.FileName <> Nothing Or OpenFileDialog1.FileName <> "") Then
            Me.Text = OpenFileDialog1.FileName.Substring(OpenFileDialog1.FileName.LastIndexOf("\") + 1, (OpenFileDialog1.FileName.IndexOf(".", 0) - (OpenFileDialog1.FileName.LastIndexOf("\") + 1))) & " - "
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim cuv() As String

        Dim tot() As Long

        Dim i As Long, j As Long

        If Len(RichTextBox1.Text) <> 0 Then

            cuv = Split(Trim(RichTextBox1.Text), " ", , vbTextCompare)

            RichTextBox2.Text = "- Cuvinte gasite: " & UBound(cuv) + 1 & vbCrLf

            ReDim tot(UBound(cuv))

            'nr aparitii

            For i = 0 To UBound(cuv)

                For j = 0 To UBound(cuv)

                    If Trim(cuv(i)) = Trim(cuv(j)) Then tot(i) = tot(i) + 1

                Next

            Next

            For i = 0 To UBound(cuv)

                For j = i + 1 To UBound(cuv)

                    If Trim(cuv(i)) = Trim(cuv(j)) Then cuv(i) = ""

                Next

            Next

            ' afisare

            For i = 0 To UBound(cuv)

                If cuv(i) <> "" Then RichTextBox2.Text = RichTextBox2.Text & "Cuvantul: " & Chr(34) & cuv(i) & Chr(34) & " nr. aparitii: " & tot(i) & vbCrLf

            Next

            RichTextBox2.Text = RichTextBox2.Text + "- "

            'afisare propozitii (succesiune de cuvinte urmate de .)

            Dim prop() As String

            prop = Split(Me.RichTextBox1.Text, ".", , vbTextCompare)

            RichTextBox2.Text = RichTextBox2.Text & "Fraze gasite: " & UBound(prop) & vbCrLf

            ReDim tot(UBound(prop))

            ' total aparitii

            For i = 0 To UBound(prop)

                For j = 0 To UBound(prop)

                    If Trim(prop(i)) = Trim(prop(j)) Then tot(i) = tot(i) + 1

                Next

            Next

            For i = 0 To UBound(prop)

                For j = i + 1 To UBound(prop)

                    If Trim(prop(i)) = Trim(prop(j)) Then prop(i) = ""

                Next

            Next

            ' afisare

            For i = 0 To UBound(prop)

                If prop(i) <> "" Then RichTextBox2.Text = RichTextBox2.Text & "Fraza: " & Chr(34) & prop(i) & Chr(34) & " nr. aparitii: " & tot(i) & vbCrLf

            Next

            RichTextBox2.Text = RichTextBox2.Text + ""

            'afisare frazele (succesiune de cuvinte urmate de . si ,)

            Dim fraz() As String

            fraz = Split(Me.RichTextBox1.Text, ".", , vbTextCompare)

            fraz = Split(Me.RichTextBox1.Text, ",", , vbTextCompare)

            'RichTextBox1.Text = RichTextBox1.Text & "frazele gasite: " & UBound(fraz) + 1 & vbCrLf

            ReDim tot(UBound(fraz))

            ' total aparitii

            For i = 0 To UBound(fraz)

                For j = 0 To UBound(fraz)

                    If Trim(fraz(i)) = Trim(fraz(j)) Then tot(i) = tot(i) + 1

                Next

            Next

            For i = 0 To UBound(fraz)

                For j = i + 1 To UBound(fraz)

                    If Trim(fraz(i)) = Trim(fraz(j)) Then fraz(i) = ""

                Next

            Next

            ' afisare

            For i = 0 To UBound(fraz)

                'If fraz(i) <> "" Then RichTextBox1.Text = RichTextBox1.Text & "frazele: " & Chr(34) & fraz(i) & Chr(34) & " nr aparitii: " & tot(i) & vbCrLf

            Next

            RichTextBox2.Text = RichTextBox2.Text + "- "

            Dim lit() As String

            ReDim lit(Len(RichTextBox1.Text))

            RichTextBox2.Text = RichTextBox2.Text & "Caractere gasite: " & UBound(lit) & vbCrLf

            ReDim tot(UBound(lit))

            For i = 0 To UBound(lit)

                lit(i) = Mid(Me.RichTextBox1.Text, i + 1, 1)

            Next

            For i = 0 To UBound(lit)

                For j = 0 To UBound(lit)

                    If lit(i) = lit(j) Then tot(i) = tot(i) + 1

                Next

            Next

            For i = 0 To UBound(lit)

                For j = i + 1 To UBound(lit)

                    If lit(i) = lit(j) Then lit(i) = ""

                Next

            Next

            ' afisare

            For i = 0 To UBound(lit)

                If lit(i) <> "" Then RichTextBox2.Text = RichTextBox2.Text & "Caracter: " & Chr(34) & lit(i) & Chr(34) & " nr aparitii: " & tot(i) & vbCrLf

            Next

        Else

            Me.RichTextBox2.Text = "Eroare: Text sursa inexistent!"

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button2.BackColor = Color.IndianRed
        Me.Hide()
        Form1.Show()

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        End
    End Sub
End Class

