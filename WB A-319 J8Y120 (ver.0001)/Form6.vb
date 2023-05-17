Imports System.Data.SqlClient
Imports System.Drawing.Printing
Imports System.Windows

Public Class Form6
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox11.Text += 1
        RichTextBox1.Clear()
        RichTextBox1.AppendText("LOADSHEET" + vbTab + vbTab + "CHECKED" + vbTab + vbTab + "APPROVED" + vbTab + vbTab + vbTab + vbTab + "EDNO" + vbNewLine)
        RichTextBox1.AppendText("ALL WEIGHTS IN KILOS" + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab + TextBox11.Text + vbNewLine)
        RichTextBox1.AppendText("FROM/TO" + vbTab + "FLIGHT" + vbTab + "A/C REG" + vbTab + vbTab + "VERSION" + vbTab + "CREW" + vbTab + "DATE" + vbTab + vbTab + "TIME" + vbNewLine)
        RichTextBox1.AppendText(Form1.TextBox3.Text + vbTab + Form1.TextBox5.Text + vbTab + Form1.TextBox4.Text + vbTab + vbTab + vbTab + Form1.TextBox2.Text + vbTab + Form1.ComboBox1.Text + vbTab + Form1.TextBox48.Text + vbTab + Form1.TextBox49.Text + vbNewLine)
        RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + "WEIGHT" + vbTab + "DISTRIBUTION" + vbNewLine)
        If TextBox1.Text = "A-319" Then
            RichTextBox1.AppendText("LOAD IN COMPARTMENTS" + vbTab + Form1.TextBox6.Text + vbTab + vbTab + "1/" + Form1.TextBox36.Text + vbTab + "4/" + Form1.TextBox37.Text + vbTab + "5/" + Form1.TextBox38.Text + vbNewLine)
        ElseIf TextBox1.Text = "A-320" Then
            RichTextBox1.AppendText("LOAD IN COMPARTMENTS" + vbTab + Form1.TextBox6.Text + vbTab + vbTab + "1/" + Form1.TextBox36.Text + vbTab + "3/" + Form1.TextBox37.Text + vbTab + "4/" + Form1.TextBox38.Text + vbTab + "5/" + Form1.TextBox16.Text + vbNewLine)
        End If
        RichTextBox1.AppendText("PASSENGER" + vbTab + vbTab + vbTab + Form1.TextBox7.Text + vbTab + vbTab + Form1.Label50.Text + "/" + Form1.Label51.Text + "/" + Form1.Label52.Text + vbTab + vbTab + vbTab + "TTL" + Form1.Label53.Text + vbNewLine)
        RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + vbTab + "PAX" + vbTab + "CY" + Form1.Label57.Text + "/" + Form1.Label56.Text + vbNewLine)
        RichTextBox1.AppendText("*************************************************************************" + vbNewLine)
        RichTextBox1.AppendText("TOTAL TRAFFIC LOAD" + vbTab + vbTab + Form1.TextBox8.Text + vbNewLine)
        RichTextBox1.AppendText("DRY OPERATING WEIGHT" + vbTab + Form1.TextBox11.Text + vbNewLine)
        RichTextBox1.AppendText("ZERO FUEL WEIGHT" + vbTab + vbTab + Form1.TextBox12.Text + vbTab + vbTab + "MAX" + Form1.TextBox13.Text + vbNewLine)
        RichTextBox1.AppendText("TAKE OFF FUEL" + vbTab + vbTab + vbTab + Form1.TextBox15.Text + vbNewLine)
        RichTextBox1.AppendText("TAKE OFF WEIGHT ACTUAL" + vbTab + Form1.TextBox17.Text + vbTab + vbTab + "MAX" + Form1.TextBox19.Text + vbNewLine)
        RichTextBox1.AppendText("TRIP FUEL" + vbTab + vbTab + vbTab + Form1.TextBox20.Text + vbNewLine)
        RichTextBox1.AppendText("LANDING WEIGHT ACTUAL" + vbTab + Form1.TextBox21.Text + vbTab + vbTab + "MAX" + Form1.TextBox18.Text + vbNewLine)
        RichTextBox1.AppendText("BALANCE AND SEATING CONDITIONS" + vbNewLine)
        RichTextBox1.AppendText("DOI" + vbTab + Form1.TextBox24.Text + vbTab + vbTab + "LIZFW" + vbTab + Form1.TextBox42.Text + vbNewLine)
        RichTextBox1.AppendText("LITOW" + vbTab + Form1.TextBox44.Text + vbTab + vbTab + "MACZFW" + Form1.TextBox43.Text + vbNewLine)
        RichTextBox1.AppendText("MACTOW" + vbTab + Form1.TextBox45.Text + vbNewLine)
        RichTextBox1.AppendText("0A/" + Form1.TextBox28.Text + vbTab + "0B/" + Form1.TextBox29.Text + vbTab + "0C/" + Form1.TextBox30.Text + vbTab + "0D/" + Form1.TextBox31.Text + vbTab + vbNewLine)
        RichTextBox1.AppendText("CABIN AREA TRIM" + vbNewLine)
        RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + vbTab + "LAST MINUTE CHANGES" + vbNewLine)
        RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + "DEST" + vbTab + "SPEC" + vbTab + "CL/CPT" + vbTab + "WEIGHT" + vbNewLine)
        If TextBox9.Text > 0 Then
            RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + TextBox22.Text + vbTab + TextBox6.Text + vbTab + TextBox7.Text + vbTab + TextBox8.Text + TextBox9.Text + vbNewLine)
        End If
        If TextBox20.Text > 0 Then
            RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + TextBox16.Text + vbTab + TextBox17.Text + vbTab + TextBox18.Text + vbTab + TextBox19.Text + TextBox20.Text + vbNewLine)
        End If
        If TextBox26.Text > 0 Then
            RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + TextBox21.Text + vbTab + TextBox23.Text + vbTab + TextBox24.Text + vbTab + TextBox25.Text + TextBox26.Text + vbNewLine)
        End If
        If TextBox31.Text > 0 Then
            RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + TextBox27.Text + vbTab + TextBox28.Text + vbTab + TextBox29.Text + vbTab + TextBox30.Text + TextBox31.Text + vbNewLine)
        End If
        If TextBox36.Text > 0 Then
            RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + TextBox32.Text + vbTab + TextBox33.Text + vbTab + TextBox34.Text + vbTab + TextBox35.Text + TextBox36.Text + vbNewLine)
        End If
        If TextBox41.Text > 0 Then
            RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + TextBox37.Text + vbTab + TextBox38.Text + vbTab + TextBox39.Text + vbTab + TextBox40.Text + TextBox41.Text + vbNewLine)
        End If
        If TextBox46.Text > 0 Then
            RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + TextBox42.Text + vbTab + TextBox43.Text + vbTab + TextBox44.Text + vbTab + TextBox45.Text + TextBox46.Text + vbNewLine)
        End If
        If TextBox53.Text > 0 Then
            RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + TextBox47.Text + vbTab + TextBox50.Text + vbTab + TextBox51.Text + vbTab + TextBox52.Text + TextBox53.Text + vbNewLine)
        End If
        If TextBox58.Text > 0 Then
            RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + TextBox54.Text + vbTab + TextBox55.Text + vbTab + TextBox56.Text + vbTab + TextBox57.Text + TextBox58.Text + vbNewLine)
        End If
        If TextBox63.Text > 0 Then
            RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + TextBox59.Text + vbTab + TextBox60.Text + vbTab + TextBox61.Text + vbTab + TextBox62.Text + TextBox63.Text + vbNewLine)
        End If
        If TextBox68.Text > 0 Then
            RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + TextBox64.Text + vbTab + TextBox65.Text + vbTab + TextBox66.Text + vbTab + TextBox67.Text + TextBox68.Text + vbNewLine)
        End If
        RichTextBox1.AppendText(vbNewLine)
        RichTextBox1.AppendText("UNDERLOAD BEFORE LMC" + Form1.Label29.Text + vbTab + vbTab + "LMC TOTAL" + TextBox10.Text + vbNewLine)
        RichTextBox1.AppendText("*************************************************************************" + vbNewLine)
        RichTextBox1.AppendText("CAPTAIN INFORMATION/NOTES" + vbNewLine)
        RichTextBox1.AppendText(vbNewLine)
        RichTextBox1.AppendText("LOAD MESSAGE BEFORE LMC" + vbNewLine)
        RichTextBox1.AppendText(vbNewLine)
        RichTextBox1.AppendText("END" + vbTab + "LOADSHEET" + vbTab + "EDNO" + TextBox11.Text + vbTab + Form1.TextBox5.Text + vbTab + Form1.TextBox48.Text + vbTab + Form1.TextBox49.Text)
    End Sub

    Private currentPage As Integer = 1 ' Текущая страница для печати

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Try
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Beep()
        End Try
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Try
            Dim font As New Font("TimesNewRoman", 12, FontStyle.Regular)
            Dim brush As Brush = Brushes.Black
            Dim marginTop As Integer = 0 ' Отступ сверху
            Dim marginLeft As Integer = 0 ' Отступ слева
            Dim printHeight As Integer = e.MarginBounds.Height
            Dim linesPerPage As Integer = printHeight \ font.Height

            Dim startIndex As Integer = (currentPage - 1) * linesPerPage
            Dim endIndex As Integer = Math.Min(startIndex + linesPerPage, RichTextBox1.Lines.Length - 1)

            Dim y As Integer = marginTop
            For i As Integer = startIndex To endIndex
                e.Graphics.DrawString(RichTextBox1.Lines(i), font, brush, marginLeft, y)
                y += font.Height
            Next i

            currentPage += 1

            If endIndex < RichTextBox1.Lines.Length - 1 Then
                e.HasMorePages = True ' Если есть еще строки для печати, устанавливаем флаг HasMorePages в True
            Else
                e.HasMorePages = False ' Если строк больше нет, устанавливаем флаг HasMorePages в False
                currentPage = 1 ' Сбрасываем счетчик страниц
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Beep()
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim wt1, wt2, wt3, wt4, wt5, wt6, wt7, wt8, wt9, wt10, wt11 As Single
        wt1 = CSng(TextBox9.Text & TextBox8.Text)
        wt2 = CSng(TextBox20.Text & TextBox19.Text)
        wt3 = CSng(TextBox26.Text & TextBox25.Text)
        wt4 = CSng(TextBox31.Text & TextBox30.Text)
        wt5 = CSng(TextBox36.Text & TextBox35.Text)
        wt6 = CSng(TextBox41.Text & TextBox40.Text)
        wt7 = CSng(TextBox46.Text & TextBox45.Text)
        wt8 = CSng(TextBox53.Text & TextBox52.Text)
        wt9 = CSng(TextBox58.Text & TextBox57.Text)
        wt10 = CSng(TextBox63.Text & TextBox62.Text)
        wt11 = CSng(TextBox68.Text & TextBox67.Text)
        TextBox10.Text = wt1 + wt2 + wt3 + wt4 + wt5 + wt6 + wt7 + wt8 + wt9 + wt10 + wt11
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form7.Show()
    End Sub
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form1.Label57.Text = Label20.Text
        Form1.Label56.Text = Label21.Text
        Form1.Label71.Text = Label71.Text
        Form1.Label73.Text = Label73.Text
        Form1.Label75.Text = Label75.Text
    End Sub
End Class