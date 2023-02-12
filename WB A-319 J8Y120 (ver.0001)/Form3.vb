Imports System.Diagnostics.Eventing.Reader
Imports System.IO.File
Public Class Form3
    Public vzr1, rb1, rm1, transvzr1, transvzr2, transvzr3 As Integer
    Public Const VZRsum As Integer = 80
    Public Const VZRwin As Integer = 85
    Public Const RBc As Integer = 30
    Public Const RMc As Integer = 15

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim VZR, RB, RM, TTL, PAX As Integer
        Dim filename As String = Application.StartupPath & "\test.log"
        Dim sw As IO.StreamWriter = AppendText(filename)
        VZR = CInt(TextBox1.Text)
        RB = CInt(TextBox2.Text)
        RM = CInt(TextBox3.Text)
        If CheckBox1.Checked Then
            TextBox4.Text = VZR * VZRsum + RB * RBc + RM * RMc
        ElseIf CheckBox2.Checked Then
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
            sw.WriteLine("[ERROR] Exceeded maximum number of passengers.")
        End If
        If PAX < 128 Then
            Form1.TextBox7.Text = TTL
            Form1.Label50.Text = VZR
            Form1.Label51.Text = RB
            Form1.Label52.Text = RM
            Form1.Label53.Text = PAX
        End If
        sw.WriteLine("[INFO] Button9_Click: VZR=" & VZR & ", RB=" & RB & ", RM=" & RM & ", PAX=" & PAX & ", TTL=" & TTL)
        sw.Close()
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
    End Sub
End Class