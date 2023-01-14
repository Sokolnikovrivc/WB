Public Class Form4
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.Clear()
        RichTextBox1.AppendText("CPM" + vbNewLine)
        RichTextBox1.AppendText(Form2.TextBox5.Text + "/" + Mid(Form2.TextBox48.Text, 1, 2) + "." + Form2.TextBox4.Text + "." + Form2.TextBox2.Text + vbNewLine)
        RichTextBox1.AppendText("-" + "11P" + "/" + Form2.TextBox7.Text + "/" + Form2.TextBox8.Text + "/" + Form2.ComboBox1.Text + "." + Form2.ComboBox4.Text + vbNewLine)
        RichTextBox1.AppendText("-" + "12P" + "/" + Form2.TextBox11.Text + "/" + Form2.TextBox12.Text + "/" + Form2.ComboBox5.Text + "." + Form2.ComboBox8.Text + vbNewLine)

    End Sub
End Class