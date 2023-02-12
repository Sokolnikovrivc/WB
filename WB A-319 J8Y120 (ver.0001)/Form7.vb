Public Class Form7
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.Clear()
        RichTextBox1.AppendText("LDM" + vbNewLine)
        RichTextBox1.AppendText(Form6.TextBox5.Text + "/" + Mid(Form6.TextBox48.Text, 1, 2) + "." + Form6.TextBox4.Text + "." + Form6.TextBox2.Text + "." + Form1.ComboBox1.Text + vbNewLine)
        RichTextBox1.AppendText("")
        RichTextBox1.AppendText("SI")
    End Sub
End Class