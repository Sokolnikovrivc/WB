Imports Microsoft.Win32

Public Class Form7
    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.Clear()
        RichTextBox1.AppendText("LDM" + vbNewLine)
        RichTextBox1.AppendText(Form6.TextBox5.Text + "/" + Mid(Form6.TextBox48.Text, 1, 2) + "." + Form6.TextBox4.Text + "." + Form6.TextBox2.Text + "." + Form1.ComboBox1.Text + vbNewLine)
        RichTextBox1.AppendText(Mid(Form6.TextBox3.Text, 1, 3) + "." + Form1.Label50.Text + "/" + Form1.Label51.Text + "/" + Form1.Label52.Text + "." + "0" + "." + "T" + Form1.TextBox6.Text + "." + "1" + "/" + Form1.TextBox36.Text + "." + "4" + "/" + Form1.TextBox37.Text + "." + "5" + "/" + Form1.TextBox38.Text + vbNewLine)
        RichTextBox1.AppendText("." + "B" + Form1.Label75.Text + "." + "C" + Form1.Label71.Text + "." + "M" + Form1.Label73.Text + vbNewLine)
        RichTextBox1.AppendText("." + "PAX" + "/" + Form1.Label57.Text + "/" + Form1.Label56.Text + "." + "PAD" + "/" + "0" + "/" + "0" + vbNewLine)
        RichTextBox1.AppendText("SI")
    End Sub
End Class