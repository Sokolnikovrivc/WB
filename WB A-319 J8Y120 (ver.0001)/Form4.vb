Public Class Form4
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Form2.TextBox1.Text = "A-319" Then
            RichTextBox1.Clear()
            RichTextBox1.AppendText("CPM" + vbNewLine)
            RichTextBox1.AppendText(Form2.TextBox5.Text + "/" + Mid(Form2.TextBox48.Text, 1, 2) + "." + Form2.TextBox4.Text + "." + Form2.TextBox2.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "11P" + "/" + Form2.TextBox7.Text + "/" + Form2.TextBox8.Text + "/" + Form2.ComboBox1.Text + "." + Form2.ComboBox4.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "12P" + "/" + Form2.TextBox11.Text + "/" + Form2.TextBox12.Text + "/" + Form2.ComboBox5.Text + "." + Form2.ComboBox8.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "41P" + "/" + Form2.TextBox22.Text + "/" + Form2.TextBox23.Text + "/" + Form2.ComboBox14.Text + "." + Form2.ComboBox13.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "42P" + "/" + Form2.TextBox19.Text + "/" + Form2.TextBox18.Text + "/" + Form2.ComboBox12.Text + "." + Form2.ComboBox9.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "51P" + "/" + Form2.TextBox31.Text + "/" + Form2.TextBox32.Text + "/" + Form2.ComboBox22.Text + "." + Form2.ComboBox21.Text + vbNewLine)
            If Form2.TextBox27.Text > 0 Then
                RichTextBox1.AppendText("-" + "51P" + "/" + Form2.TextBox28.Text + "/" + Form2.TextBox27.Text + "/" + Form2.ComboBox20.Text + "." + Form2.ComboBox17.Text + vbNewLine)
            End If
            RichTextBox1.AppendText("SI")
        End If
        If Form9.TextBox1.Text = "A-320" Then
            RichTextBox1.Clear()
            RichTextBox1.AppendText("CPM" + vbNewLine)
            RichTextBox1.AppendText(Form9.TextBox5.Text + "/" + Mid(Form9.TextBox48.Text, 1, 2) + "." + Form9.TextBox4.Text + "." + Form9.TextBox2.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "11P" + "/" + Form9.TextBox7.Text + "/" + Form9.TextBox8.Text + "/" + Form9.ComboBox1.Text + "." + Form9.ComboBox4.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "12P" + "/" + Form9.TextBox11.Text + "/" + Form9.TextBox12.Text + "/" + Form9.ComboBox5.Text + "." + Form9.ComboBox8.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "13P" + "/" + Form9.TextBox15.Text + "/" + Form9.TextBox16.Text + "/" + Form9.ComboBox9.Text + "." + Form9.ComboBox12.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "31P" + "/" + Form9.TextBox25.Text + "/" + Form9.TextBox26.Text + "/" + Form9.ComboBox18.Text + "." + Form9.ComboBox17.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "32P" + "/" + Form9.TextBox22.Text + "/" + Form9.TextBox21.Text + "/" + Form9.ComboBox16.Text + "." + Form9.ComboBox13.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "41P" + "/" + Form9.TextBox35.Text + "/" + Form9.TextBox36.Text + "/" + Form9.ComboBox26.Text + "." + Form9.ComboBox25.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "42P" + "/" + Form9.TextBox31.Text + "/" + Form9.TextBox30.Text + "/" + Form9.ComboBox24.Text + "." + Form9.ComboBox21.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "51P" + "/" + Form9.TextBox53.Text + "/" + Form9.TextBox45.Text + "/" + Form9.ComboBox34.Text + "." + Form9.ComboBox33.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "52P" + "/" + Form9.TextBox41.Text + "/" + Form9.TextBox40.Text + "/" + Form9.ComboBox32.Text + "." + Form9.ComboBox29.Text + vbNewLine)
            RichTextBox1.AppendText("-" + "53P" + "/" + Form9.TextBox47.Text + "/" + Form9.TextBox50.Text + "/" + Form9.ComboBox37.Text + "." + Form9.ComboBox40.Text + vbNewLine)
            RichTextBox1.AppendText("SI")
        End If
    End Sub
End Class