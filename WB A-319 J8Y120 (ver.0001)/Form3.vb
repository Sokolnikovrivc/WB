Public Class Form3
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim VZR, RB, RM, TTL, PAX As Integer
        Const VZRsum As Integer = 80
        Const VZRwin As Integer = 85
        Const RBc As Integer = 30
        Const RMc As Integer = 15
        VZR = CInt(TextBox1.Text)
        RB = CInt(TextBox2.Text)
        RM = CInt(TextBox3.Text)
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False
        ElseIf CheckBox2.Checked = True Then
            CheckBox1.Checked = False
        End If
        If CheckBox1.Checked Then
            TextBox4.Text = VZR * VZRsum + RB * RBc + RM * RMc
        End If
        If CheckBox2.Checked Then
            TextBox4.Text = VZR * VZRwin + RB * RBc + RM * RMc
        End If
        TextBox5.Text = VZR + RB + RM
        TTL = CSng(TextBox4.Text)
        PAX = CInt(TextBox5.Text)
        If PAX > 128 Then
            MsgBox("Превышение максимального значения пассажиров!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
        End If
        Form1.TextBox7.Text = TTL
        Form1.Label50.Text = VZR
        Form1.Label51.Text = RB
        Form1.Label52.Text = RM
        Form1.Label53.Text = PAX
    End Sub
End Class