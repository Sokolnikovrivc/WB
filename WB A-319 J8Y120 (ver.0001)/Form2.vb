
Imports System.Data.SqlClient
Imports System.Drawing.Printing
Imports System.Net.Security
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button

Public Class Form2
    Public wt1, wt2, wt3, wt4, wt5, wt6 As Single


    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim f As Font = RichTextBox1.Font
        Dim b As Brush = Brushes.Black
        Dim x As Integer = e.MarginBounds.Left
        Dim y As Integer = e.MarginBounds.Top
        Dim linesPerPage As Integer = e.MarginBounds.Height \ f.Height
        Dim lineCount As Integer = 0
        Dim text As String = ""
        'Получаем все строки из RichTextBox, которые помещаются на текущей странице
        While lineCount < linesPerPage AndAlso RichTextBox1.Lines.Length > 0
            text = RichTextBox1.Lines(0)
            RichTextBox1.Lines = RichTextBox1.Lines.Skip(1).ToArray()
            lineCount += 1

            'Проверяем, помещается ли текущая строка на текущей странице
            If e.Graphics.MeasureString(text, f).Width > e.MarginBounds.Width Then
                'Если строка не помещается на текущей странице, то переносим ее на следующую страницу
                RichTextBox1.Lines = New String() {text}.Concat(RichTextBox1.Lines).ToArray()
                Exit While
            End If

            'Печатаем текущую строку на текущей странице
            e.Graphics.DrawString(text, f, b, x, y)
            y += f.Height
        End While

        'Если в RichTextBox остались строки, которые не помещаются на текущей странице, то печатаем их на следующей странице
        If RichTextBox1.Lines.Length > 0 Then
            e.HasMorePages = True
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If PrintDialog1.ShowDialog() = DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form4.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim MAX11 As Integer = registryKey.GetValue("TextBox19", "")
        Dim MAX12 As Integer = registryKey.GetValue("TextBox20", "")
        Dim MAX41 As Integer = registryKey.GetValue("TextBox21", "")
        Dim MAX42 As Integer = registryKey.GetValue("TextBox22", "")
        Dim MAX51 As Integer = registryKey.GetValue("TextBox23", "")
        Dim MAXV11 As Single = registryKey.GetValue("TextBox24", "")
        Dim MAXV12 As Single = registryKey.GetValue("TextBox25", "")
        Dim MAXV41 As Single = registryKey.GetValue("TextBox26", "")
        Dim MAXV42 As Single = registryKey.GetValue("TextBox27", "")
        Dim MAXV51 As Single = registryKey.GetValue("TextBox28", "")
        registryKey.Close()
        Dim Mwt11 As Single
        Dim Mwt12 As Single
        Dim V11 As Single
        Dim V12 As Single
        Dim Mwt41 As Single
        Dim Mwt42 As Single
        Dim V41 As Single
        Dim V42 As Single
        Dim Mwt51a As Single
        Dim Mwt51b As Single
        Dim V51a As Single
        Dim V51b As Single
        Dim a, b, c, d, k, f, g, h, i, t, r, w As Single
        'позиции 11 и 12
        If ComboBox4.Text = "RFL" And ComboBox8.Text = "ROX" Or ComboBox4.Text = "ROX" And ComboBox8.Text = "RFL" Then
            MsgBox("Несовместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RSC" And ComboBox8.Text = "ROX" Or ComboBox4.Text = "ROX" And ComboBox8.Text = "RSC" Then
            MsgBox("Несовместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RFW" And ComboBox8.Text = "RCM" Or ComboBox4.Text = "RCM" And ComboBox8.Text = "RFW" Then
            MsgBox("Несовместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And ComboBox8.Text = "RPB" Or ComboBox4.Text = "RPB" And ComboBox8.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And ComboBox8.Text = "RIS" Or ComboBox4.Text = "RIS" And ComboBox8.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And ComboBox8.Text = "ICE" Or ComboBox4.Text = "ICE" And ComboBox8.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RPB" And ComboBox8.Text = "HEG" Or ComboBox4.Text = "HEG" And ComboBox8.Text = "RPB" Then
            MsgBox("Несовместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RIS" And ComboBox8.Text = "HEG" Or ComboBox4.Text = "HEG" And ComboBox8.Text = "RIS" Then
            MsgBox("Несовместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RPB" And TextBox9.Text = "PEM" Or TextBox9.Text = "RPB" And ComboBox4.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RPB" And TextBox9.Text = "PEP" Or TextBox9.Text = "RPB" And ComboBox4.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RPB" And TextBox9.Text = "PES" Or TextBox9.Text = "RPB" And ComboBox4.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RPB" And TextBox9.Text = "EAT" Or TextBox9.Text = "RPB" And ComboBox4.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RIS" And TextBox9.Text = "PEM" Or TextBox9.Text = "RIS" And ComboBox4.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RIS" And TextBox9.Text = "PEP" Or TextBox9.Text = "RIS" And ComboBox4.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RIS" And TextBox9.Text = "PES" Or TextBox9.Text = "RIS" And ComboBox4.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RIS" And TextBox9.Text = "EAT" Or TextBox9.Text = "RIS" And ComboBox4.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And TextBox9.Text = "AVI" Or TextBox9.Text = "AVI" And ComboBox4.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And TextBox9.Text = "HUM" Or TextBox9.Text = "AVI" And ComboBox4.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "HUM" And TextBox9.Text = "PEM" Or TextBox9.Text = "HUM" And ComboBox4.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "HUM" And TextBox9.Text = "PEP" Or TextBox9.Text = "HUM" And ComboBox4.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "HUM" And TextBox9.Text = "PES" Or TextBox9.Text = "HUM" And ComboBox4.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "HUM" And TextBox9.Text = "EAT" Or TextBox9.Text = "HUM" And ComboBox4.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RPB" And TextBox14.Text = "PEM" Or TextBox14.Text = "RPB" And ComboBox8.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RPB" And TextBox14.Text = "PEP" Or TextBox14.Text = "RPB" And ComboBox8.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RPB" And TextBox14.Text = "PES" Or TextBox14.Text = "RPB" And ComboBox8.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RPB" And TextBox14.Text = "EAT" Or TextBox14.Text = "RPB" And ComboBox8.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RIS" And TextBox14.Text = "PEM" Or TextBox14.Text = "RIS" And ComboBox8.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RIS" And TextBox14.Text = "PEP" Or TextBox14.Text = "RIS" And ComboBox8.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RIS" And TextBox14.Text = "PES" Or TextBox14.Text = "RIS" And ComboBox8.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RIS" And TextBox14.Text = "EAT" Or TextBox14.Text = "RIS" And ComboBox8.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "AVI" And TextBox14.Text = "AVI" Or TextBox14.Text = "AVI" And ComboBox8.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "AVI" And TextBox14.Text = "HUM" Or TextBox14.Text = "AVI" And ComboBox8.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "HUM" And TextBox14.Text = "PEM" Or TextBox14.Text = "HUM" And ComboBox8.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "HUM" And TextBox14.Text = "PEP" Or TextBox14.Text = "HUM" And ComboBox8.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "HUM" And TextBox14.Text = "PES" Or TextBox14.Text = "HUM" And ComboBox8.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "HUM" And TextBox14.Text = "EAT" Or TextBox14.Text = "HUM" And ComboBox8.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        Mwt11 = CSng(TextBox8.Text)
        Mwt12 = CSng(TextBox12.Text)
        V11 = CSng(TextBox10.Text)
        V12 = CSng(TextBox13.Text)
        If Mwt11 > MAX11 Then
            MsgBox("Превышение максимального веса в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If Mwt12 > MAX12 Then
            MsgBox("Превышение максимального веса в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V11 > MAXV11 Then
            MsgBox("Превышение максимального объёма в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V12 > MAXV12 Then
            MsgBox("Превышение максимального объёма в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        'позиции 41 и 42
        If ComboBox13.Text = "RFL" And ComboBox9.Text = "ROX" Or ComboBox13.Text = "ROX" And ComboBox9.Text = "RFL" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RSC" And ComboBox9.Text = "ROX" Or ComboBox13.Text = "ROX" And ComboBox9.Text = "RSC" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RFW" And ComboBox9.Text = "RCM" Or ComboBox13.Text = "RCM" And ComboBox9.Text = "RFW" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "AVI" And ComboBox9.Text = "RPB" Or ComboBox13.Text = "RPB" And ComboBox9.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "AVI" And ComboBox9.Text = "RIS" Or ComboBox13.Text = "RIS" And ComboBox9.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "AVI" And ComboBox9.Text = "ICE" Or ComboBox13.Text = "ICE" And ComboBox9.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RPB" And ComboBox9.Text = "HEG" Or ComboBox13.Text = "HEG" And ComboBox9.Text = "RPB" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RIS" And ComboBox9.Text = "HEG" Or ComboBox13.Text = "HEG" And ComboBox9.Text = "RIS" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RPB" And TextBox20.Text = "PEM" Or TextBox20.Text = "RPB" And ComboBox13.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RPB" And TextBox20.Text = "PEP" Or TextBox20.Text = "RPB" And ComboBox13.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RPB" And TextBox20.Text = "PES" Or TextBox20.Text = "RPB" And ComboBox13.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RPB" And TextBox20.Text = "EAT" Or TextBox20.Text = "RPB" And ComboBox13.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RIS" And TextBox20.Text = "PEM" Or TextBox20.Text = "RIS" And ComboBox13.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RIS" And TextBox20.Text = "PEP" Or TextBox20.Text = "RIS" And ComboBox13.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RIS" And TextBox20.Text = "PES" Or TextBox20.Text = "RIS" And ComboBox13.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RIS" And TextBox20.Text = "EAT" Or TextBox20.Text = "RIS" And ComboBox13.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "AVI" And TextBox20.Text = "AVI" Or TextBox20.Text = "AVI" And ComboBox13.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "AVI" And TextBox20.Text = "HUM" Or TextBox20.Text = "AVI" And ComboBox13.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "HUM" And TextBox20.Text = "PEM" Or TextBox20.Text = "HUM" And ComboBox13.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "HUM" And TextBox20.Text = "PEP" Or TextBox20.Text = "HUM" And ComboBox13.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "HUM" And TextBox20.Text = "PES" Or TextBox20.Text = "HUM" And ComboBox13.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "HUM" And TextBox20.Text = "EAT" Or TextBox20.Text = "HUM" And ComboBox13.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "RPB" And TextBox16.Text = "PEM" Or TextBox16.Text = "RPB" And ComboBox9.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "RPB" And TextBox16.Text = "PEP" Or TextBox16.Text = "RPB" And ComboBox9.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "RPB" And TextBox16.Text = "PES" Or TextBox16.Text = "RPB" And ComboBox9.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "RPB" And TextBox16.Text = "EAT" Or TextBox16.Text = "RPB" And ComboBox9.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "RIS" And TextBox16.Text = "PEM" Or TextBox16.Text = "RIS" And ComboBox9.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "RIS" And TextBox16.Text = "PEP" Or TextBox16.Text = "RIS" And ComboBox9.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "RIS" And TextBox16.Text = "PES" Or TextBox16.Text = "RIS" And ComboBox9.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "RIS" And TextBox16.Text = "EAT" Or TextBox16.Text = "RIS" And ComboBox9.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "AVI" And TextBox16.Text = "AVI" Or TextBox16.Text = "AVI" And ComboBox9.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "AVI" And TextBox16.Text = "HUM" Or TextBox16.Text = "AVI" And ComboBox9.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "HUM" And TextBox16.Text = "PEM" Or TextBox16.Text = "HUM" And ComboBox9.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "HUM" And TextBox16.Text = "PEP" Or TextBox16.Text = "HUM" And ComboBox9.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "HUM" And TextBox16.Text = "PES" Or TextBox16.Text = "HUM" And ComboBox9.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox9.Text = "HUM" And TextBox16.Text = "EAT" Or TextBox16.Text = "HUM" And ComboBox9.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        Mwt41 = CSng(TextBox23.Text)
        Mwt42 = CSng(TextBox18.Text)
        V41 = CSng(TextBox21.Text)
        V42 = CSng(TextBox17.Text)
        If Mwt41 > MAX41 Then
            MsgBox("Превышение максимального веса в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If Mwt42 > MAX42 Then
            MsgBox("Превышение максимального веса в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V41 > MAXV41 Then
            MsgBox("Превышение максимального объёма в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V42 > MAXV42 Then
            MsgBox("Превышение максимального объёма в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        'позиции 51
        If ComboBox21.Text = "RFL" And ComboBox17.Text = "ROX" Or ComboBox17.Text = "ROX" And ComboBox21.Text = "RFL" Then
            MsgBox("Несовместимые грузы в позициях 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RSC" And ComboBox17.Text = "ROX" Or ComboBox17.Text = "ROX" And ComboBox21.Text = "RSC" Then
            MsgBox("Несовместимые грузы в позициях 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RFW" And ComboBox17.Text = "RCM" Or ComboBox17.Text = "RCM" And ComboBox21.Text = "RFW" Then
            MsgBox("Несовместимые грузы в позициях 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "AVI" And ComboBox17.Text = "RPB" Or ComboBox17.Text = "RPB" And ComboBox21.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "AVI" And ComboBox17.Text = "RIS" Or ComboBox17.Text = "RIS" And ComboBox21.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "AVI" And ComboBox17.Text = "ICE" Or ComboBox17.Text = "ICE" And ComboBox21.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RPB" And ComboBox17.Text = "HEG" Or ComboBox17.Text = "HEG" And ComboBox21.Text = "RPB" Then
            MsgBox("Несовместимые грузы в позициях 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RIS" And ComboBox17.Text = "HEG" Or ComboBox17.Text = "HEG" And ComboBox21.Text = "RIS" Then
            MsgBox("Несовместимые грузы в позициях 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RPB" And TextBox29.Text = "PEM" Or TextBox29.Text = "RPB" And ComboBox21.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RPB" And TextBox29.Text = "PEP" Or TextBox29.Text = "RPB" And ComboBox21.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RPB" And TextBox29.Text = "PES" Or TextBox29.Text = "RPB" And ComboBox21.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RPB" And TextBox29.Text = "EAT" Or TextBox29.Text = "RPB" And ComboBox21.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RIS" And TextBox29.Text = "PEM" Or TextBox29.Text = "RIS" And ComboBox21.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RIS" And TextBox29.Text = "PEP" Or TextBox29.Text = "RIS" And ComboBox21.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RIS" And TextBox29.Text = "PES" Or TextBox29.Text = "RIS" And ComboBox21.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox24.Text = "RIS" And TextBox29.Text = "EAT" Or TextBox29.Text = "RIS" And ComboBox21.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox24.Text = "AVI" And TextBox29.Text = "AVI" Or TextBox29.Text = "AVI" And ComboBox21.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox24.Text = "AVI" And TextBox29.Text = "HUM" Or TextBox29.Text = "AVI" And ComboBox21.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox24.Text = "HUM" And TextBox29.Text = "PEM" Or TextBox29.Text = "HUM" And ComboBox21.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox24.Text = "HUM" And TextBox29.Text = "PEP" Or TextBox29.Text = "HUM" And ComboBox21.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox24.Text = "HUM" And TextBox29.Text = "PES" Or TextBox29.Text = "HUM" And ComboBox21.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox24.Text = "HUM" And TextBox29.Text = "EAT" Or TextBox29.Text = "HUM" And ComboBox21.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RPB" And TextBox25.Text = "PEM" Or TextBox25.Text = "RPB" And ComboBox17.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RPB" And TextBox25.Text = "PEP" Or TextBox25.Text = "RPB" And ComboBox17.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RPB" And TextBox25.Text = "PES" Or TextBox25.Text = "RPB" And ComboBox17.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RPB" And TextBox25.Text = "EAT" Or TextBox25.Text = "RPB" And ComboBox17.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RIS" And TextBox25.Text = "PEM" Or TextBox25.Text = "RIS" And ComboBox17.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RIS" And TextBox25.Text = "PEP" Or TextBox25.Text = "RIS" And ComboBox17.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RIS" And TextBox25.Text = "PES" Or TextBox25.Text = "RIS" And ComboBox17.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RIS" And TextBox25.Text = "EAT" Or TextBox25.Text = "RIS" And ComboBox17.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "AVI" And TextBox25.Text = "AVI" Or TextBox25.Text = "AVI" And ComboBox17.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "AVI" And TextBox25.Text = "HUM" Or TextBox25.Text = "AVI" And ComboBox17.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "HUM" And TextBox25.Text = "PEM" Or TextBox25.Text = "HUM" And ComboBox17.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "HUM" And TextBox25.Text = "PEP" Or TextBox25.Text = "HUM" And ComboBox17.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "HUM" And TextBox25.Text = "PES" Or TextBox25.Text = "HUM" And ComboBox17.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "HUM" And TextBox25.Text = "EAT" Or TextBox25.Text = "HUM" And ComboBox17.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        Mwt51a = CSng(TextBox32.Text)
        Mwt51b = CSng(TextBox27.Text)
        V51a = CSng(TextBox30.Text)
        V51b = CSng(TextBox26.Text)
        If Mwt51a + Mwt51b > MAX51 Then
            MsgBox("Превышение максимального веса в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V51a + V51b > MAXV51 Then
            MsgBox("Превышение максимального объёма в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        wt1 = CSng(TextBox8.Text)
        wt2 = CSng(TextBox12.Text)
        wt3 = CSng(TextBox23.Text)
        wt4 = CSng(TextBox18.Text)
        wt5 = CSng(TextBox32.Text)
        wt6 = CSng(TextBox27.Text)
        'ДЛЯ ПОЗИЦИЙ 11 И 12
        If ComboBox1.Text = "F" And ComboBox5.Text = "F" Then
            Label45.Text = wt1 + wt2
        ElseIf ComboBox1.Text = "F" Then
            Label45.Text = wt1
        ElseIf ComboBox5.Text = "F" Then
            Label45.Text = wt2
        End If
        If ComboBox1.Text = "C" And ComboBox5.Text = "C" Then
            Label46.Text = wt1 + wt2
        ElseIf ComboBox1.Text = "C" Then
            Label46.Text = wt1
        ElseIf ComboBox5.Text = "C" Then
            Label46.Text = wt2
        End If
        If ComboBox1.Text = "M" And ComboBox5.Text = "M" Then
            Label48.Text = wt1 + wt2
        ElseIf ComboBox1.Text = "M" Then
            Label48.Text = wt1
        ElseIf ComboBox5.Text = "M" Then
            Label48.Text = wt2
        End If
        If ComboBox1.Text = "B" And ComboBox5.Text = "B" Then
            Label50.Text = wt1 + wt2
        ElseIf ComboBox1.Text = "B" Then
            Label50.Text = wt1
        ElseIf ComboBox5.Text = "B" Then
            Label50.Text = wt2
        End If
        'ДЛЯ ПОЗИЦИЙ 41 И 42
        If ComboBox14.Text = "F" And ComboBox12.Text = "F" Then
            Label58.Text = wt3 + wt4
        ElseIf ComboBox14.Text = "F" Then
            Label58.Text = wt3
        ElseIf ComboBox12.Text = "F" Then
            Label58.Text = wt4
        End If
        If ComboBox14.Text = "C" And ComboBox12.Text = "C" Then
            Label56.Text = wt3 + wt4
        ElseIf ComboBox14.Text = "C" Then
            Label56.Text = wt3
        ElseIf ComboBox12.Text = "C" Then
            Label56.Text = wt4
        End If
        If ComboBox14.Text = "M" And ComboBox12.Text = "M" Then
            Label54.Text = wt3 + wt4
        ElseIf ComboBox14.Text = "M" Then
            Label54.Text = wt3
        ElseIf ComboBox12.Text = "M" Then
            Label54.Text = wt4
        End If
        If ComboBox14.Text = "B" And ComboBox12.Text = "B" Then
            Label52.Text = wt3 + wt4
        ElseIf ComboBox14.Text = "B" Then
            Label52.Text = wt3
        ElseIf ComboBox12.Text = "B" Then
            Label52.Text = wt4
        End If
        'ДЛЯ ПОЗИЦИЙ 51
        If ComboBox22.Text = "F" And ComboBox20.Text = "F" Then
            Label66.Text = wt5 + wt6
        ElseIf ComboBox22.Text = "F" Then
            Label66.Text = wt5
        ElseIf ComboBox20.Text = "F" Then
            Label66.Text = wt6
        End If
        If ComboBox22.Text = "C" And ComboBox20.Text = "C" Then
            Label64.Text = wt5 + wt6
        ElseIf ComboBox22.Text = "C" Then
            Label64.Text = wt5
        ElseIf ComboBox20.Text = "C" Then
            Label64.Text = wt6
        End If
        If ComboBox22.Text = "M" And ComboBox20.Text = "M" Then
            Label62.Text = wt5 + wt6
        ElseIf ComboBox22.Text = "M" Then
            Label62.Text = wt5
        ElseIf ComboBox20.Text = "M" Then
            Label62.Text = wt6
        End If
        If ComboBox22.Text = "B" And ComboBox20.Text = "B" Then
            Label60.Text = wt5 + wt6
        ElseIf ComboBox22.Text = "B" Then
            Label60.Text = wt5
        ElseIf ComboBox20.Text = "B" Then
            Label60.Text = wt6
        End If
        a = CSng(Label45.Text)
        b = CSng(Label58.Text)
        c = CSng(Label66.Text)
        Label69.Text = a + b + c
        k = CSng(Label46.Text)
        d = CSng(Label56.Text)
        f = CSng(Label64.Text)
        Label71.Text = k + d + f
        g = CSng(Label48.Text)
        h = CSng(Label54.Text)
        i = CSng(Label62.Text)
        Label73.Text = g + h + i
        t = CSng(Label50.Text)
        r = CSng(Label52.Text)
        w = CSng(Label60.Text)
        Label75.Text = t + r + w
        TextBox6.Text = wt1 + wt2
        TextBox15.Text = wt3 + wt4
        TextBox24.Text = wt5 + wt6
        Form1.Label71.Text = Label71.Text
        Form1.Label73.Text = Label73.Text
        Form1.Label75.Text = Label75.Text
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox33.Text += 1
        Dim connection As New SqlConnection("Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True")
        Dim command As New SqlCommand("Insert into [Test].[dbo].[A319 lir] ([flight_id1],[flight_route],[flight_bort],[config_id],[date_flight],[time_flight],[CPT1FWD],[dest11],[LIC11] ,[type11],[wt11],[stat11],[code11],[V11],[info11],[dest12],[LIC12],[type12],[wt12],[stat12],[code12],[V12],[info12],[CPT4AFT],[dest41],[LIC41],[type41],[wt41],[stat41],[code41],[V41],[info41],[dest42],[LIC42],[type42],[wt42],[stat42],[code42],[V42],[info42],[CPT5AFT],[dest51],[LIC51],[type51],[wt51],[stat51],[code51],[V51],[info51],[dest52],[LIC52],[type52],[wt52],[stat52],[code52],[V52],[info52]) Values(@flight_id1, @flight_route, @flight_bort, @config_id, @date_flight, @time_flight, @CPT1FWD, @dest11, @LIC11, @type11, @wt11, @stat11, @code11, @V11, @info11, @dest12, @LIC12, @type12, @wt12, @stat12, @code12, @V12, @info12, @CPT4AFT, @dest41, @LIC41, @type41, @wt41, @stat41, @code41, @V41, @info41, @dest42, @LIC42 , @type42 , @wt42 , @stat42 , @code42 , @V42, @info42, @CPT5AFT, @dest51, @LIC51 , @type51 , @wt51 , @stat51 , @code51 , @V51 , @info51 , @dest52 , @LIC52 , @type52 , @wt52 , @stat52, @code52 , @V52 , @info52)", connection)
        command.Parameters.AddWithValue("@flight_id1", TextBox5.Text)
        command.Parameters.AddWithValue("@flight_route", TextBox3.Text)
        command.Parameters.AddWithValue("@flight_bort", TextBox4.Text)
        command.Parameters.AddWithValue("@config_id", TextBox2.Text)
        command.Parameters.AddWithValue("@date_flight", TextBox48.Text)
        command.Parameters.AddWithValue("@time_flight", TextBox49.Text)
        command.Parameters.AddWithValue("@CPT1FWD", TextBox6.Text)
        command.Parameters.AddWithValue("@dest11", TextBox7.Text)
        command.Parameters.AddWithValue("@LIC11", ComboBox1.Text)
        command.Parameters.AddWithValue("@type11", ComboBox2.Text)
        command.Parameters.AddWithValue("@wt11", TextBox8.Text)
        command.Parameters.AddWithValue("@stat11", ComboBox3.Text)
        command.Parameters.AddWithValue("@code11", ComboBox4.Text)
        command.Parameters.AddWithValue("@V11", TextBox10.Text)
        command.Parameters.AddWithValue("@info11", TextBox9.Text)
        command.Parameters.AddWithValue("@dest12", TextBox11.Text)
        command.Parameters.AddWithValue("@LIC12", ComboBox5.Text)
        command.Parameters.AddWithValue("@type12", ComboBox6.Text)
        command.Parameters.AddWithValue("@wt12", TextBox12.Text)
        command.Parameters.AddWithValue("@stat12", ComboBox7.Text)
        command.Parameters.AddWithValue("@code12", ComboBox8.Text)
        command.Parameters.AddWithValue("@V12", TextBox13.Text)
        command.Parameters.AddWithValue("@info12", TextBox14.Text)
        command.Parameters.AddWithValue("@CPT4AFT", TextBox15.Text)
        command.Parameters.AddWithValue("@dest41", TextBox22.Text)
        command.Parameters.AddWithValue("@LIC41", ComboBox14.Text)
        command.Parameters.AddWithValue("@type41", ComboBox15.Text)
        command.Parameters.AddWithValue("@wt41", TextBox23.Text)
        command.Parameters.AddWithValue("@stat41", ComboBox16.Text)
        command.Parameters.AddWithValue("@code41", ComboBox13.Text)
        command.Parameters.AddWithValue("@V41", TextBox21.Text)
        command.Parameters.AddWithValue("@info41", TextBox20.Text)
        command.Parameters.AddWithValue("@dest42", TextBox19.Text)
        command.Parameters.AddWithValue("@LIC42", ComboBox12.Text)
        command.Parameters.AddWithValue("@type42", ComboBox11.Text)
        command.Parameters.AddWithValue("@wt42", TextBox18.Text)
        command.Parameters.AddWithValue("@stat42", ComboBox10.Text)
        command.Parameters.AddWithValue("@code42", ComboBox9.Text)
        command.Parameters.AddWithValue("@V42", TextBox17.Text)
        command.Parameters.AddWithValue("@info42", TextBox16.Text)
        command.Parameters.AddWithValue("@CPT5AFT", TextBox24.Text)
        command.Parameters.AddWithValue("@dest51", TextBox31.Text)
        command.Parameters.AddWithValue("@LIC51", ComboBox22.Text)
        command.Parameters.AddWithValue("@type51", ComboBox23.Text)
        command.Parameters.AddWithValue("@wt51", TextBox32.Text)
        command.Parameters.AddWithValue("@stat51", ComboBox24.Text)
        command.Parameters.AddWithValue("@code51", ComboBox21.Text)
        command.Parameters.AddWithValue("@V51", TextBox30.Text)
        command.Parameters.AddWithValue("@info51", TextBox29.Text)
        command.Parameters.AddWithValue("@dest52", TextBox28.Text)
        command.Parameters.AddWithValue("@LIC52", ComboBox20.Text)
        command.Parameters.AddWithValue("@type52", ComboBox19.Text)
        command.Parameters.AddWithValue("@wt52", TextBox27.Text)
        command.Parameters.AddWithValue("@stat52", ComboBox18.Text)
        command.Parameters.AddWithValue("@code52", ComboBox17.Text)
        command.Parameters.AddWithValue("@V52", TextBox26.Text)
        command.Parameters.AddWithValue("@info52", TextBox25.Text)
        connection.Open()
        command.ExecuteNonQuery()
        connection.Close()
        MessageBox.Show("Запись в Базу Данных")
        RichTextBox1.Clear()
        RichTextBox1.AppendText("LOADING INSTRUCTION/REPORT" + vbTab + vbTab + "PREPARED BY" + vbTab + vbTab + vbTab + "EDNO" + TextBox33.Text + vbNewLine)
        RichTextBox1.AppendText("ALL WEIGHTS IN KILOS" + vbNewLine)
        RichTextBox1.AppendText("FROM/TO" + vbTab + "FLIGHT" + vbTab + "A/C REG" + vbTab + "VERSION" + vbTab + "CREW" + vbTab + "DATE" + vbTab + vbTab + vbTab + "TIME" + vbNewLine)
        RichTextBox1.AppendText(TextBox3.Text + vbTab + TextBox5.Text + vbTab + TextBox4.Text + vbTab + vbTab + TextBox2.Text + vbTab + vbTab + Form1.ComboBox1.Text + vbTab + TextBox48.Text + vbTab + vbTab + TextBox49.Text + vbNewLine)
        RichTextBox1.AppendText("PLANNED JOINING LOAD" + vbNewLine)
        RichTextBox1.AppendText(TextBox7.Text + vbTab + vbTab + "F" + Label69.Text + vbTab + vbTab + vbTab + "C" + Label71.Text + vbTab + vbTab + "M" + Label73.Text + vbTab + vbTab + "B" + Label75.Text + vbNewLine)
        RichTextBox1.AppendText("JOINING SPECS : SEE SUMMARY" + vbNewLine)
        RichTextBox1.AppendText("TRANSIT SPECS: SEE SUMMARY" + vbNewLine)
        RichTextBox1.AppendText("LOADING INSTRUCTIONS" + vbNewLine)
        RichTextBox1.AppendText("******************************************************************************************************************************" + vbNewLine)
        RichTextBox1.AppendText("CPT1FWD" + vbTab + "MAX" + "2268" + vbNewLine)
        RichTextBox1.AppendText(":" + "11P" + vbTab + vbTab + ComboBox2.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox3.Text + ":" + vbTab + TextBox7.Text + vbTab + ComboBox1.Text + "/" + TextBox8.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox4.Text + vbTab + TextBox9.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox8.Text + vbNewLine)
        RichTextBox1.AppendText("......................................................................................................" + vbNewLine)
        RichTextBox1.AppendText(":" + "12P" + vbTab + vbTab + ComboBox6.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox7.Text + ":" + vbTab + TextBox11.Text + vbTab + ComboBox5.Text + "/" + TextBox12.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox8.Text + vbTab + TextBox14.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox12.Text + vbTab + vbTab + vbTab + "CPT1FWD" + vbTab + "TTL" + vbTab + TextBox6.Text + vbNewLine)
        RichTextBox1.AppendText("******************************************************************************************************************************" + vbNewLine)
        RichTextBox1.AppendText("CPT4AFT" + vbTab + vbTab + "MAX" + "3021" + vbNewLine)
        RichTextBox1.AppendText(":" + "41P" + vbTab + vbTab + ComboBox15.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox16.Text + ":" + vbTab + TextBox22.Text + vbTab + ComboBox14.Text + "/" + TextBox23.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox13.Text + vbTab + TextBox20.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox23.Text + vbNewLine)
        RichTextBox1.AppendText("......................................................................................................" + vbNewLine)
        RichTextBox1.AppendText(":" + "42P" + vbTab + vbTab + ComboBox11.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox10.Text + ":" + vbTab + TextBox19.Text + vbTab + ComboBox12.Text + "/" + TextBox18.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox9.Text + vbTab + TextBox16.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox18.Text + vbTab + vbTab + vbTab + "CPT4AFT" + vbTab + vbTab + "TTL" + vbTab + TextBox15.Text + vbNewLine)
        RichTextBox1.AppendText("******************************************************************************************************************************" + vbNewLine)
        RichTextBox1.AppendText("CPT5AFT" + vbTab + vbTab + "MAX" + "1497" + vbNewLine)
        RichTextBox1.AppendText(":" + "51P" + vbTab + vbTab + ComboBox23.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox24.Text + ":" + vbTab + TextBox31.Text + vbTab + ComboBox22.Text + "/" + TextBox32.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox21.Text + vbTab + TextBox29.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox32.Text + vbNewLine)
        If TextBox27.Text > 0 Then
            RichTextBox1.AppendText("......................................................................................................" + vbNewLine)
            RichTextBox1.AppendText(":" + "51P" + vbTab + vbTab + ComboBox19.Text + vbNewLine)
            RichTextBox1.AppendText(":" + ComboBox18.Text + ":" + vbTab + TextBox28.Text + vbTab + ComboBox20.Text + "/" + TextBox27.Text + vbNewLine)
            RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox17.Text + vbTab + TextBox25.Text + vbNewLine)
        End If
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox27.Text + vbTab + vbTab + vbTab + "CPT5AFT" + vbTab + vbTab + "TTL" + vbTab + TextBox24.Text + vbNewLine)
        RichTextBox1.AppendText("******************************************************************************************************************************" + vbNewLine)
        RichTextBox1.AppendText(":" + "SI" + vbNewLine)
        RichTextBox1.AppendText("THIS AIRCRAFT HAS BEEN LOADED IN ACCORDANCE WITH THESE INSTRUCTIONS AND THE DEVIATIONS" + vbNewLine)
        RichTextBox1.AppendText("SHOWN ON THIS REPORT. THE CONTAINERS / PALLETS AND BULKLOAD HAVE BEEN SECURED IN" + vbNewLine)
        RichTextBox1.AppendText("ACCORDANCE WITH COMPANY INSTRUCTIONS." + vbNewLine)
        RichTextBox1.AppendText(vbNewLine)
        RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + "SIGNATURE")
        Form1.TextBox36.Text = TextBox6.Text
        Form1.TextBox37.Text = TextBox15.Text
        Form1.TextBox38.Text = TextBox24.Text
        Form1.TextBox6.Text = wt1 + wt2 + wt3 + wt4 + wt5 + wt6
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label45.Text = "0"
        Label46.Text = "0"
        Label48.Text = "0"
        Label50.Text = "0"
        Label58.Text = "0"
        Label56.Text = "0"
        Label54.Text = "0"
        Label52.Text = "0"
        Label66.Text = "0"
        Label64.Text = "0"
        Label62.Text = "0"
        Label60.Text = "0"
        Label69.Text = "0"
        Label71.Text = "0"
        Label73.Text = "0"
        Label75.Text = "0"
        TextBox8.Text = "0"
        TextBox12.Text = "0"
        TextBox23.Text = "0"
        TextBox18.Text = "0"
        TextBox32.Text = "0"
        TextBox27.Text = "0"
        TextBox6.Text = "0"
        TextBox15.Text = "0"
        TextBox24.Text = "0"
    End Sub
    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Form2Data")
        ' Save the form data to the registry
        registryKey.SetValue("TextBox8", TextBox8.Text)
        registryKey.SetValue("TextBox12", TextBox12.Text)
        registryKey.Close()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Form2Data")
        ' Load the form data from the registry
        TextBox8.Text = registryKey.GetValue("TextBox8", "")
        TextBox12.Text = registryKey.GetValue("TextBox12", "")
        registryKey.Close()
        If (String.IsNullOrEmpty(TextBox8.Text)) Then
            TextBox8.Text = "0"
        End If
        If (String.IsNullOrEmpty(TextBox12.Text)) Then
            TextBox12.Text = "0"
        End If
    End Sub
End Class