Imports System.Diagnostics.Eventing.Reader
Imports System.IO.File
Public Class Form3
    Public vzr1, rb1, rm1, transvzr1, transvzr2, transvzr3 As Integer
    Public vzr2, rb2, rm2, transvzr4, transvzr5, transvzr6 As Integer
    Public vzr3, rb3, rm3, transvzr7, transvzr8, transvzr9 As Integer
    Public vzr4, rb4, rm4, transvzr10, transvzr11, transvzr12, ttl1, ttl2, ttl3, ttl4 As Integer
    Public Const VZRsum As Integer = 80
    Public Const VZRwin As Integer = 85
    Public Const RBc As Integer = 30
    Public Const RMc As Integer = 15

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim VZR, RB, RM, TTL, PAX As Integer
        Dim filename As String = Application.StartupPath & "\test.log"
        Dim sw As IO.StreamWriter = AppendText(filename)
        TextBox1.Text = vzr1 + vzr2 + vzr3 + vzr4 + transvzr1 + transvzr4 + transvzr7 + transvzr10
        TextBox2.Text = rb1 + rb2 + rb3 + rb4 + transvzr2 + transvzr5 + transvzr8 + transvzr11
        TextBox3.Text = rm1 + rm2 + rm3 + rm4 + transvzr3 + transvzr6 + transvzr9 + transvzr12
        TextBox4.Text = ttl1 + ttl2 + ttl3 + ttl4
        TextBox7.Text = transvzr1 + transvzr4 + transvzr7 + transvzr10 + transvzr2 + transvzr5 + transvzr8 + transvzr11 + transvzr3 + transvzr6 + transvzr9 + transvzr12
        VZR = CInt(TextBox1.Text)
        RB = CInt(TextBox2.Text)
        RM = CInt(TextBox3.Text)
        TTL = CInt(TextBox4.Text)
        TextBox5.Text = VZR + RB + RM
        PAX = CInt(TextBox5.Text)
        If PAX > 128 Then
            MsgBox("Превышение максимального значения пассажиров!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            sw.WriteLine("[ERROR] Exceeded maximum number of passengers.")
        End If
        If PAX < 128 Then
            Form1.TextBox7.Text = TTL
            Form1.Label50.Text = VZR
            Form1.Label51.Text = RB
            Form1.Label52.Text = RM
            Form1.Label53.Text = PAX
            Form1.TextBox28.Text = ttl1
            Form1.TextBox29.Text = ttl2
            Form1.TextBox30.Text = ttl3
            Form1.TextBox31.Text = ttl4
        End If
        sw.WriteLine("[INFO] Button9_Click: VZR=" & VZR & ", RB=" & RB & ", RM=" & RM & ", PAX=" & PAX & ", TTL=" & TTL)
        sw.Close()
        If ComboBox1.Text = "F" Or ComboBox1.Text = "J" Or ComboBox1.Text = "C" Then
            Form6.Label20.Text = vzr1 + rb1 + rm1 + transvzr1 + transvzr2 + transvzr3
        ElseIf ComboBox1.Text = "Y" And ComboBox2.Text = "Y" And ComboBox3.Text = "Y" And ComboBox4.Text = "Y" Then
            Form6.Label21.Text = vzr1 + rb1 + rm1 + transvzr1 + transvzr2 + transvzr3 + vzr2 + rb2 + rm2 + transvzr4 + transvzr5 + transvzr6 + vzr3 + rb3 + rm3 + transvzr7 + transvzr8 + transvzr9 + vzr4 + rb4 + rm4 + transvzr10 + transvzr11 + transvzr12
        End If
        If ComboBox2.Text = "Y" And ComboBox3.Text = "Y" And ComboBox4.Text = "Y" Then
            Form6.Label21.Text = vzr2 + rb2 + rm2 + transvzr4 + transvzr5 + transvzr6 + vzr3 + rb3 + rm3 + transvzr7 + transvzr8 + transvzr9 + vzr4 + rb4 + rm4 + transvzr10 + transvzr11 + transvzr12
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False
        ElseIf CheckBox2.Checked = True Then
            CheckBox1.Checked = False
        End If
        GroupBox1.Visible = CheckBox1.Checked
        GroupBox2.Visible = CheckBox1.Checked
        GroupBox3.Visible = CheckBox1.Checked
        GroupBox4.Visible = CheckBox1.Checked
        Button1.Visible = CheckBox1.Checked
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False
        ElseIf CheckBox1.Checked = True Then
            CheckBox2.Checked = False
        End If
        GroupBox1.Visible = CheckBox2.Checked
        GroupBox2.Visible = CheckBox2.Checked
        GroupBox3.Visible = CheckBox2.Checked
        GroupBox4.Visible = CheckBox2.Checked
        Button1.Visible = CheckBox2.Checked
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        vzr1 = CInt(TextBox8.Text)
        rb1 = CInt(TextBox9.Text)
        rm1 = CInt(TextBox10.Text)
        transvzr1 = CInt(TextBox11.Text)
        transvzr2 = CInt(TextBox12.Text)
        transvzr3 = CInt(TextBox13.Text)
        vzr2 = CInt(TextBox19.Text)
        rb2 = CInt(TextBox18.Text)
        rm2 = CInt(TextBox17.Text)
        transvzr4 = CInt(TextBox16.Text)
        transvzr5 = CInt(TextBox15.Text)
        transvzr6 = CInt(TextBox14.Text)
        vzr3 = CInt(TextBox25.Text)
        rb3 = CInt(TextBox24.Text)
        rm3 = CInt(TextBox23.Text)
        transvzr7 = CInt(TextBox22.Text)
        transvzr8 = CInt(TextBox21.Text)
        transvzr9 = CInt(TextBox20.Text)
        vzr4 = CInt(TextBox31.Text)
        rb4 = CInt(TextBox30.Text)
        rm4 = CInt(TextBox29.Text)
        transvzr10 = CInt(TextBox28.Text)
        transvzr11 = CInt(TextBox27.Text)
        transvzr12 = CInt(TextBox26.Text)
        If vzr1 + rb1 + rm1 + transvzr1 + transvzr2 + transvzr3 > 8 Then
            MsgBox("Превышение максимального значения пассажиров!,Зона 0А", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox8.Text = "0"
            TextBox9.Text = "0"
            TextBox10.Text = "0"
            TextBox11.Text = "0"
            TextBox12.Text = "0"
            TextBox13.Text = "0"
            Label33.Text = "0"
        End If
        If CheckBox1.Checked Then
            Label33.Text = (vzr1 * VZRsum + transvzr1 * VZRsum) + (rb1 * RBc + transvzr2 * RBc) + (rm1 * RMc + transvzr3 * RMc)
        ElseIf CheckBox2.Checked Then
            Label33.Text = (vzr1 * VZRwin + transvzr1 * VZRsum) + (rb1 * RBc + transvzr2 * RBc) + (rm1 * RMc + transvzr3 * RMc)
        End If
        If vzr2 + rb2 + rm2 + transvzr4 + transvzr5 + transvzr6 > 42 Then
            MsgBox("Превышение максимального значения пассажиров!,Зона 0B", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox19.Text = "0"
            TextBox18.Text = "0"
            TextBox17.Text = "0"
            TextBox16.Text = "0"
            TextBox15.Text = "0"
            TextBox14.Text = "0"
            Label14.Text = "0"
        End If
        If CheckBox1.Checked Then
            Label14.Text = (vzr2 * VZRsum + transvzr4 * VZRsum) + (rb2 * RBc + transvzr5 * RBc) + (rm2 * RMc + transvzr6 * RMc)
        ElseIf CheckBox2.Checked Then
            Label14.Text = (vzr2 * VZRwin + transvzr4 * VZRsum) + (rb2 * RBc + transvzr5 * RBc) + (rm2 * RMc + transvzr6 * RMc)
        End If
        If vzr3 + rb3 + rm3 + transvzr7 + transvzr8 + transvzr9 > 42 Then
            MsgBox("Превышение максимального значения пассажиров!,Зона 0C", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox25.Text = "0"
            TextBox24.Text = "0"
            TextBox23.Text = "0"
            TextBox22.Text = "0"
            TextBox21.Text = "0"
            TextBox20.Text = "0"
            Label34.Text = "0"
        End If
        If CheckBox1.Checked Then
            Label34.Text = (vzr3 * VZRsum + transvzr7 * VZRsum) + (rb3 * RBc + transvzr8 * RBc) + (rm3 * RMc + transvzr9 * RMc)
        ElseIf CheckBox2.Checked Then
            Label34.Text = (vzr3 * VZRwin + transvzr7 * VZRsum) + (rb3 * RBc + transvzr8 * RBc) + (rm3 * RMc + transvzr9 * RMc)
        End If
        If vzr4 + rb4 + rm4 + transvzr10 + transvzr11 + transvzr12 > 36 Then
            MsgBox("Превышение максимального значения пассажиров!,Зона 0D", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox31.Text = "0"
            TextBox30.Text = "0"
            TextBox29.Text = "0"
            TextBox28.Text = "0"
            TextBox27.Text = "0"
            TextBox26.Text = "0"
            Label37.Text = "0"
        End If
        If CheckBox1.Checked Then
            Label37.Text = (vzr4 * VZRsum + transvzr10 * VZRsum) + (rb4 * RBc + transvzr11 * RBc) + (rm4 * RMc + transvzr12 * RMc)
        ElseIf CheckBox2.Checked Then
            Label37.Text = (vzr4 * VZRwin + transvzr10 * VZRsum) + (rb4 * RBc + transvzr11 * RBc) + (rm4 * RMc + transvzr12 * RMc)
        End If
        ttl1 = CSng(Label33.Text)
        ttl2 = CSng(Label14.Text)
        ttl3 = CSng(Label34.Text)
        ttl4 = CSng(Label37.Text)
    End Sub
End Class