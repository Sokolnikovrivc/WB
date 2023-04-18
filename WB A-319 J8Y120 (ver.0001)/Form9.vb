Imports System.Data.SqlClient
Imports System.Drawing.Printing
Public Class Form9
    Public wt1, wt2, wt3, wt4, wt5, wt6, wt7, wt8, wt9, wt10 As Single

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form4.Show()
    End Sub
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

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Mwt11 As Single
        Dim Mwt12 As Single
        Dim Mwt13 As Single
        Dim V11 As Single
        Dim V12 As Single
        Dim V13 As Single
        Dim Mwt31 As Single
        Dim Mwt32 As Single
        Dim V31 As Single
        Dim V32 As Single
        Dim Mwt41 As Single
        Dim Mwt42 As Single
        Dim V41 As Single
        Dim V42 As Single
        Dim Mwt51 As Single
        Dim Mwt52 As Single
        Dim Mwt53 As Single
        Dim V51 As Single
        Dim V52 As Single
        Dim V53 As Single
        Dim a, b, c, d, k, f, g, h, i, t, r, w, z, l, q, y As Single
        'позиции 11 и 12, 13
        If ComboBox4.Text = "RFL" And ComboBox8.Text = "ROX" Or ComboBox4.Text = "ROX" And ComboBox8.Text = "RFL" Or ComboBox12.Text = "ROX" And ComboBox4.Text = "RFL" Or ComboBox12.Text = "RFL" And ComboBox4.Text = "ROX" Or ComboBox12.Text = "ROX" And ComboBox8.Text = "RFL" Or ComboBox12.Text = "RFL" And ComboBox8.Text = "ROX" Then
            MsgBox("Несовместимые грузы в позициях 11,12,13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RSC" And ComboBox8.Text = "ROX" Or ComboBox4.Text = "ROX" And ComboBox8.Text = "RSC" Or ComboBox12.Text = "RSC" And ComboBox4.Text = "ROX" Or ComboBox12.Text = "ROX" And ComboBox4.Text = "RSC" Or ComboBox12.Text = "RSC" And ComboBox8.Text = "ROX" Or ComboBox12.Text = "ROX" And ComboBox8.Text = "RSC" Then
            MsgBox("Несовместимые грузы в позициях 11,12,13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RFW" And ComboBox8.Text = "RCM" Or ComboBox4.Text = "RCM" And ComboBox8.Text = "RFW" Or ComboBox12.Text = "RFW" And ComboBox4.Text = "RCM" Or ComboBox12.Text = "RCM" And ComboBox4.Text = "RFW" Or ComboBox12.Text = "RFW" And ComboBox8.Text = "RCM" Or ComboBox12.Text = "RCM" And ComboBox8.Text = "RFW" Then
            MsgBox("Несовместимые грузы в позициях 11,12,13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And ComboBox8.Text = "RPB" Or ComboBox4.Text = "RPB" And ComboBox8.Text = "AVI" Or ComboBox12.Text = "AVI" And ComboBox4.Text = "RPB" Or ComboBox12.Text = "RPB" And ComboBox4.Text = "AVI" Or ComboBox12.Text = "AVI" And ComboBox8.Text = "RPB" Or ComboBox12.Text = "RPB" And ComboBox8.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 11,12,13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And ComboBox8.Text = "RIS" Or ComboBox4.Text = "RIS" And ComboBox8.Text = "AVI" Or ComboBox12.Text = "AVI" And ComboBox4.Text = "RIS" Or ComboBox12.Text = "RIS" And ComboBox4.Text = "AVI" Or ComboBox12.Text = "AVI" And ComboBox8.Text = "RIS" Or ComboBox12.Text = "RIS" And ComboBox8.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 11,12,13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And ComboBox8.Text = "ICE" Or ComboBox4.Text = "ICE" And ComboBox8.Text = "AVI" Or ComboBox12.Text = "AVI" And ComboBox4.Text = "ICE" Or ComboBox12.Text = "ICE" And ComboBox4.Text = "AVI" Or ComboBox12.Text = "AVI" And ComboBox8.Text = "ICE" Or ComboBox12.Text = "ICE" And ComboBox8.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 11,12,13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RPB" And ComboBox8.Text = "HEG" Or ComboBox4.Text = "HEG" And ComboBox8.Text = "RPB" Or ComboBox12.Text = "RPB" And ComboBox4.Text = "HEG" Or ComboBox12.Text = "HEG" And ComboBox4.Text = "RPB" Or ComboBox12.Text = "RPB" And ComboBox8.Text = "HEG" Or ComboBox12.Text = "HEG" And ComboBox8.Text = "RPB" Then
            MsgBox("Несовместимые грузы в позициях 11,12,13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RIS" And ComboBox8.Text = "HEG" Or ComboBox4.Text = "HEG" And ComboBox8.Text = "RIS" Or ComboBox12.Text = "RIS" And ComboBox4.Text = "HEG" Or ComboBox12.Text = "HEG" And ComboBox4.Text = "RIS" Or ComboBox12.Text = "RIS" And ComboBox8.Text = "HEG" Or ComboBox12.Text = "HEG" And ComboBox8.Text = "RIS" Then
            MsgBox("Несовместимые грузы в позициях 11,12,13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
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
        If ComboBox12.Text = "RPB" And TextBox18.Text = "PEM" Or TextBox18.Text = "RPB" And ComboBox12.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "RPB" And TextBox18.Text = "PEP" Or TextBox18.Text = "RPB" And ComboBox12.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "RPB" And TextBox18.Text = "PES" Or TextBox18.Text = "RPB" And ComboBox12.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "RPB" And TextBox18.Text = "EAT" Or TextBox18.Text = "RPB" And ComboBox12.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "RIS" And TextBox18.Text = "PEM" Or TextBox18.Text = "RIS" And ComboBox12.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "RIS" And TextBox18.Text = "PEP" Or TextBox18.Text = "RIS" And ComboBox12.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "RIS" And TextBox18.Text = "PES" Or TextBox18.Text = "RIS" And ComboBox12.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "RIS" And TextBox18.Text = "EAT" Or TextBox18.Text = "RIS" And ComboBox12.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "AVI" And TextBox18.Text = "AVI" Or TextBox18.Text = "AVI" And ComboBox12.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "AVI" And TextBox18.Text = "HUM" Or TextBox18.Text = "AVI" And ComboBox12.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "HUM" And TextBox18.Text = "PEM" Or TextBox18.Text = "HUM" And ComboBox12.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "HUM" And TextBox18.Text = "PEP" Or TextBox18.Text = "HUM" And ComboBox12.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "HUM" And TextBox18.Text = "PES" Or TextBox18.Text = "HUM" And ComboBox12.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox12.Text = "HUM" And TextBox18.Text = "EAT" Or TextBox18.Text = "HUM" And ComboBox12.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        Mwt11 = CSng(TextBox8.Text)
        Mwt12 = CSng(TextBox12.Text)
        Mwt13 = CSng(TextBox16.Text)
        V11 = CSng(TextBox10.Text)
        V12 = CSng(TextBox13.Text)
        V13 = CSng(TextBox17.Text)
        If Mwt11 > 1045 Then
            MsgBox("Превышение максимального веса в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If Mwt12 > 1225 Then
            MsgBox("Превышение максимального веса в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If Mwt13 > 1132 Then
            MsgBox("Превышение максимального веса в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V11 > 4.09 Then
            MsgBox("Превышение максимального объёма в позиции 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V12 > 4.77 Then
            MsgBox("Превышение максимального объёма в позиции 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V13 > 4.32 Then
            MsgBox("Превышение максимального объёма в позиции 13!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        '31 & 32
        If ComboBox13.Text = "RFL" And ComboBox17.Text = "ROX" Or ComboBox13.Text = "ROX" And ComboBox17.Text = "RFL" Then
            MsgBox("Несовместимые грузы в позициях 31 и 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RSC" And ComboBox17.Text = "ROX" Or ComboBox13.Text = "ROX" And ComboBox17.Text = "RSC" Then
            MsgBox("Несовместимые грузы в позициях 31 и 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RFW" And ComboBox17.Text = "RCM" Or ComboBox13.Text = "RCM" And ComboBox17.Text = "RFW" Then
            MsgBox("Несовместимые грузы в позициях 31 и 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "AVI" And ComboBox17.Text = "RPB" Or ComboBox13.Text = "RPB" And ComboBox17.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 31 и 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "AVI" And ComboBox17.Text = "RIS" Or ComboBox13.Text = "RIS" And ComboBox17.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 31 и 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "AVI" And ComboBox17.Text = "ICE" Or ComboBox13.Text = "ICE" And ComboBox17.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 31 и 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RPB" And ComboBox17.Text = "HEG" Or ComboBox13.Text = "HEG" And ComboBox17.Text = "RPB" Then
            MsgBox("Несовместимые грузы в позициях 31 и 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RIS" And ComboBox17.Text = "HEG" Or ComboBox13.Text = "HEG" And ComboBox17.Text = "RIS" Then
            MsgBox("Несовместимые грузы в позициях 31 и 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RPB" And ComboBox19.Text = "PEM" Or ComboBox19.Text = "RPB" And ComboBox13.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RPB" And ComboBox19.Text = "PEP" Or ComboBox19.Text = "RPB" And ComboBox13.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RPB" And ComboBox19.Text = "PES" Or ComboBox19.Text = "RPB" And ComboBox13.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RPB" And ComboBox19.Text = "EAT" Or ComboBox19.Text = "RPB" And ComboBox13.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RIS" And ComboBox19.Text = "PEM" Or ComboBox19.Text = "RIS" And ComboBox13.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RIS" And ComboBox19.Text = "PEP" Or ComboBox19.Text = "RIS" And ComboBox13.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RIS" And ComboBox19.Text = "PES" Or ComboBox19.Text = "RIS" And ComboBox13.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "RIS" And ComboBox19.Text = "EAT" Or ComboBox19.Text = "RIS" And ComboBox13.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "AVI" And ComboBox19.Text = "AVI" Or ComboBox19.Text = "AVI" And ComboBox13.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "AVI" And ComboBox19.Text = "HUM" Or ComboBox19.Text = "AVI" And ComboBox13.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "HUM" And ComboBox19.Text = "PEM" Or ComboBox19.Text = "HUM" And ComboBox13.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "HUM" And ComboBox19.Text = "PEP" Or ComboBox19.Text = "HUM" And ComboBox13.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "HUM" And ComboBox19.Text = "PES" Or ComboBox19.Text = "HUM" And ComboBox13.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox13.Text = "HUM" And ComboBox19.Text = "EAT" Or ComboBox19.Text = "HUM" And ComboBox13.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RPB" And TextBox23.Text = "PEM" Or TextBox23.Text = "RPB" And ComboBox17.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RPB" And TextBox23.Text = "PEP" Or TextBox23.Text = "RPB" And ComboBox17.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RPB" And TextBox23.Text = "PES" Or TextBox23.Text = "RPB" And ComboBox17.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RPB" And TextBox23.Text = "EAT" Or TextBox23.Text = "RPB" And ComboBox17.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RIS" And TextBox23.Text = "PEM" Or TextBox23.Text = "RIS" And ComboBox17.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RIS" And TextBox23.Text = "PEP" Or TextBox23.Text = "RIS" And ComboBox17.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RIS" And TextBox23.Text = "PES" Or TextBox23.Text = "RIS" And ComboBox17.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "RIS" And TextBox23.Text = "EAT" Or TextBox23.Text = "RIS" And ComboBox17.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "AVI" And TextBox23.Text = "AVI" Or TextBox23.Text = "AVI" And ComboBox17.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "AVI" And TextBox23.Text = "HUM" Or TextBox23.Text = "AVI" And ComboBox17.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "HUM" And TextBox23.Text = "PEM" Or TextBox23.Text = "HUM" And ComboBox17.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "HUM" And TextBox23.Text = "PEP" Or TextBox23.Text = "HUM" And ComboBox17.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "HUM" And TextBox23.Text = "PES" Or TextBox23.Text = "HUM" And ComboBox17.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox17.Text = "HUM" And TextBox23.Text = "EAT" Or TextBox23.Text = "HUM" And ComboBox17.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        Mwt31 = CSng(TextBox26.Text)
        Mwt32 = CSng(TextBox21.Text)
        V31 = CSng(TextBox24.Text)
        V32 = CSng(TextBox20.Text)
        If Mwt31 > 1301 Then
            MsgBox("Превышение максимального веса в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If Mwt32 > 1125 Then
            MsgBox("Превышение максимального веса в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V31 > 5.23 Then
            MsgBox("Превышение максимального объёма в позиции 31!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V32 > 4.53 Then
            MsgBox("Превышение максимального объёма в позиции 32!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        '41 и 42 
        If ComboBox25.Text = "RFL" And ComboBox21.Text = "ROX" Or ComboBox21.Text = "ROX" And ComboBox25.Text = "RFL" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RSC" And ComboBox21.Text = "ROX" Or ComboBox21.Text = "ROX" And ComboBox25.Text = "RSC" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RFW" And ComboBox21.Text = "RCM" Or ComboBox21.Text = "RCM" And ComboBox25.Text = "RFW" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "AVI" And ComboBox21.Text = "RPB" Or ComboBox21.Text = "RPB" And ComboBox25.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "AVI" And ComboBox21.Text = "RIS" Or ComboBox21.Text = "RIS" And ComboBox25.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "AVI" And ComboBox21.Text = "ICE" Or ComboBox21.Text = "ICE" And ComboBox25.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RPB" And ComboBox21.Text = "HEG" Or ComboBox21.Text = "HEG" And ComboBox25.Text = "RPB" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RIS" And ComboBox21.Text = "HEG" Or ComboBox21.Text = "HEG" And ComboBox25.Text = "RIS" Then
            MsgBox("Несовместимые грузы в позициях 41 и 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RPB" And TextBox32.Text = "PEM" Or TextBox32.Text = "RPB" And ComboBox25.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RPB" And TextBox32.Text = "PEP" Or TextBox32.Text = "RPB" And ComboBox25.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RPB" And TextBox32.Text = "PES" Or TextBox32.Text = "RPB" And ComboBox25.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RPB" And TextBox32.Text = "EAT" Or TextBox32.Text = "RPB" And ComboBox25.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RIS" And TextBox32.Text = "PEM" Or TextBox32.Text = "RIS" And ComboBox25.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RIS" And TextBox32.Text = "PEP" Or TextBox32.Text = "RIS" And ComboBox25.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RIS" And TextBox32.Text = "PES" Or TextBox32.Text = "RIS" And ComboBox25.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "RIS" And TextBox32.Text = "EAT" Or TextBox32.Text = "RIS" And ComboBox25.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "AVI" And TextBox32.Text = "AVI" Or TextBox32.Text = "AVI" And ComboBox25.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "AVI" And TextBox32.Text = "HUM" Or TextBox32.Text = "AVI" And ComboBox25.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "HUM" And TextBox32.Text = "PEM" Or TextBox32.Text = "HUM" And ComboBox25.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "HUM" And TextBox32.Text = "PEP" Or TextBox32.Text = "HUM" And ComboBox25.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "HUM" And TextBox32.Text = "PES" Or TextBox32.Text = "HUM" And ComboBox25.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox25.Text = "HUM" And TextBox32.Text = "EAT" Or TextBox32.Text = "HUM" And ComboBox25.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RPB" And TextBox28.Text = "PEM" Or TextBox28.Text = "RPB" And ComboBox21.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RPB" And TextBox28.Text = "PEP" Or TextBox28.Text = "RPB" And ComboBox21.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RPB" And TextBox28.Text = "PES" Or TextBox28.Text = "RPB" And ComboBox21.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RPB" And TextBox28.Text = "EAT" Or TextBox28.Text = "RPB" And ComboBox21.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RIS" And TextBox28.Text = "PEM" Or TextBox28.Text = "RIS" And ComboBox21.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RIS" And TextBox28.Text = "PEP" Or TextBox28.Text = "RIS" And ComboBox21.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RIS" And TextBox28.Text = "PES" Or TextBox28.Text = "RIS" And ComboBox21.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "RIS" And TextBox28.Text = "EAT" Or TextBox28.Text = "RIS" And ComboBox21.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "AVI" And TextBox28.Text = "AVI" Or TextBox28.Text = "AVI" And ComboBox21.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "AVI" And TextBox28.Text = "HUM" Or TextBox28.Text = "AVI" And ComboBox21.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "HUM" And TextBox28.Text = "PEM" Or TextBox28.Text = "HUM" And ComboBox21.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "HUM" And TextBox28.Text = "PEP" Or TextBox28.Text = "HUM" And ComboBox21.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "HUM" And TextBox28.Text = "PES" Or TextBox28.Text = "HUM" And ComboBox21.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox21.Text = "HUM" And TextBox28.Text = "EAT" Or TextBox28.Text = "HUM" And ComboBox21.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        Mwt41 = CSng(TextBox36.Text)
        Mwt42 = CSng(TextBox30.Text)
        V41 = CSng(TextBox34.Text)
        V42 = CSng(TextBox29.Text)
        If Mwt41 > 928 Then
            MsgBox("Превышение максимального веса в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If Mwt42 > 1182 Then
            MsgBox("Превышение максимального веса в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V41 > 3.75 Then
            MsgBox("Превышение максимального объёма в позиции 41!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V42 > 4.75 Then
            MsgBox("Превышение максимального объёма в позиции 42!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        'Bulk 
        If ComboBox33.Text = "RFL" And ComboBox29.Text = "ROX" Or ComboBox33.Text = "ROX" And ComboBox29.Text = "RFL" Or ComboBox40.Text = "ROX" And ComboBox33.Text = "RFL" Or ComboBox40.Text = "RFL" And ComboBox33.Text = "ROX" Or ComboBox40.Text = "ROX" And ComboBox29.Text = "RFL" Or ComboBox40.Text = "RFL" And ComboBox29.Text = "ROX" Then
            MsgBox("Несовместимые грузы в позициях 51,52,53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RSC" And ComboBox29.Text = "ROX" Or ComboBox33.Text = "ROX" And ComboBox29.Text = "RSC" Or ComboBox40.Text = "RSC" And ComboBox33.Text = "ROX" Or ComboBox40.Text = "ROX" And ComboBox33.Text = "RSC" Or ComboBox40.Text = "RSC" And ComboBox29.Text = "ROX" Or ComboBox40.Text = "ROX" And ComboBox29.Text = "RSC" Then
            MsgBox("Несовместимые грузы в позициях 51,52,53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RFW" And ComboBox29.Text = "RCM" Or ComboBox33.Text = "RCM" And ComboBox29.Text = "RFW" Or ComboBox40.Text = "RFW" And ComboBox33.Text = "RCM" Or ComboBox40.Text = "RCM" And ComboBox33.Text = "RFW" Or ComboBox40.Text = "RFW" And ComboBox29.Text = "RCM" Or ComboBox40.Text = "RCM" And ComboBox29.Text = "RFW" Then
            MsgBox("Несовместимые грузы в позициях 51,52,53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "AVI" And ComboBox29.Text = "RPB" Or ComboBox33.Text = "RPB" And ComboBox29.Text = "AVI" Or ComboBox40.Text = "AVI" And ComboBox33.Text = "RPB" Or ComboBox40.Text = "RPB" And ComboBox33.Text = "AVI" Or ComboBox40.Text = "AVI" And ComboBox29.Text = "RPB" Or ComboBox40.Text = "RPB" And ComboBox29.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 51,52,53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "AVI" And ComboBox29.Text = "RIS" Or ComboBox33.Text = "RIS" And ComboBox29.Text = "AVI" Or ComboBox40.Text = "AVI" And ComboBox33.Text = "RIS" Or ComboBox40.Text = "RIS" And ComboBox33.Text = "AVI" Or ComboBox40.Text = "AVI" And ComboBox29.Text = "RIS" Or ComboBox40.Text = "RIS" And ComboBox29.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 51,52,53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "AVI" And ComboBox29.Text = "ICE" Or ComboBox33.Text = "ICE" And ComboBox29.Text = "AVI" Or ComboBox40.Text = "AVI" And ComboBox33.Text = "ICE" Or ComboBox40.Text = "ICE" And ComboBox33.Text = "AVI" Or ComboBox40.Text = "AVI" And ComboBox29.Text = "ICE" Or ComboBox40.Text = "ICE" And ComboBox29.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позициях 51,52,53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RPB" And ComboBox29.Text = "HEG" Or ComboBox33.Text = "HEG" And ComboBox29.Text = "RPB" Or ComboBox40.Text = "RPB" And ComboBox33.Text = "HEG" Or ComboBox40.Text = "HEG" And ComboBox33.Text = "RPB" Or ComboBox40.Text = "RPB" And ComboBox29.Text = "HEG" Or ComboBox40.Text = "HEG" And ComboBox29.Text = "RPB" Then
            MsgBox("Несовместимые грузы в позициях 51,52,53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RIS" And ComboBox29.Text = "HEG" Or ComboBox33.Text = "HEG" And ComboBox29.Text = "RIS" Or ComboBox40.Text = "RIS" And ComboBox33.Text = "HEG" Or ComboBox40.Text = "HEG" And ComboBox33.Text = "RIS" Or ComboBox40.Text = "RIS" And ComboBox29.Text = "HEG" Or ComboBox40.Text = "HEG" And ComboBox29.Text = "RIS" Then
            MsgBox("Несовместимые грузы в позициях 51,52,53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RPB" And TextBox42.Text = "PEM" Or TextBox42.Text = "RPB" And ComboBox33.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RPB" And TextBox42.Text = "PEP" Or TextBox42.Text = "RPB" And ComboBox33.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RPB" And TextBox42.Text = "PES" Or TextBox42.Text = "RPB" And ComboBox33.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RPB" And TextBox42.Text = "EAT" Or TextBox42.Text = "RPB" And ComboBox33.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RIS" And TextBox42.Text = "PEM" Or TextBox42.Text = "RIS" And ComboBox33.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RIS" And TextBox42.Text = "PEP" Or TextBox42.Text = "RIS" And ComboBox33.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RIS" And TextBox42.Text = "PES" Or TextBox42.Text = "RIS" And ComboBox33.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "RIS" And TextBox42.Text = "EAT" Or TextBox42.Text = "RIS" And ComboBox33.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "AVI" And TextBox42.Text = "AVI" Or TextBox42.Text = "AVI" And ComboBox33.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "AVI" And TextBox42.Text = "HUM" Or TextBox42.Text = "AVI" And ComboBox33.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "HUM" And TextBox42.Text = "PEM" Or TextBox42.Text = "HUM" And ComboBox33.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "HUM" And TextBox42.Text = "PEP" Or TextBox42.Text = "HUM" And ComboBox33.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "HUM" And TextBox42.Text = "PES" Or TextBox42.Text = "HUM" And ComboBox33.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox33.Text = "HUM" And TextBox42.Text = "EAT" Or TextBox42.Text = "HUM" And ComboBox33.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "RPB" And TextBox38.Text = "PEM" Or TextBox38.Text = "RPB" And ComboBox29.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "RPB" And TextBox38.Text = "PEP" Or TextBox38.Text = "RPB" And ComboBox29.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "RPB" And TextBox38.Text = "PES" Or TextBox38.Text = "RPB" And ComboBox29.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "RPB" And TextBox38.Text = "EAT" Or TextBox38.Text = "RPB" And ComboBox29.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "RIS" And TextBox38.Text = "PEM" Or TextBox38.Text = "RIS" And ComboBox29.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "RIS" And TextBox38.Text = "PEP" Or TextBox38.Text = "RIS" And ComboBox29.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "RIS" And TextBox38.Text = "PES" Or TextBox38.Text = "RIS" And ComboBox29.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "RIS" And TextBox38.Text = "EAT" Or TextBox38.Text = "RIS" And ComboBox29.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "AVI" And TextBox38.Text = "AVI" Or TextBox38.Text = "AVI" And ComboBox29.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "AVI" And TextBox38.Text = "HUM" Or TextBox38.Text = "AVI" And ComboBox29.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "HUM" And TextBox38.Text = "PEM" Or TextBox38.Text = "HUM" And ComboBox29.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "HUM" And TextBox38.Text = "PEP" Or TextBox38.Text = "HUM" And ComboBox29.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "HUM" And TextBox38.Text = "PES" Or TextBox38.Text = "HUM" And ComboBox29.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox29.Text = "HUM" And TextBox38.Text = "EAT" Or TextBox38.Text = "HUM" And ComboBox29.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "RPB" And TextBox52.Text = "PEM" Or TextBox52.Text = "RPB" And ComboBox40.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "RPB" And TextBox52.Text = "PEP" Or TextBox52.Text = "RPB" And ComboBox40.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "RPB" And TextBox52.Text = "PES" Or TextBox52.Text = "RPB" And ComboBox40.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "RPB" And TextBox52.Text = "EAT" Or TextBox52.Text = "RPB" And ComboBox40.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "RIS" And TextBox52.Text = "PEM" Or TextBox52.Text = "RIS" And ComboBox40.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "RIS" And TextBox52.Text = "PEP" Or TextBox52.Text = "RIS" And ComboBox40.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "RIS" And TextBox52.Text = "PES" Or TextBox52.Text = "RIS" And ComboBox40.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "RIS" And TextBox52.Text = "EAT" Or TextBox52.Text = "RIS" And ComboBox40.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "AVI" And TextBox52.Text = "AVI" Or TextBox52.Text = "AVI" And ComboBox40.Text = "AVI" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "AVI" And TextBox52.Text = "HUM" Or TextBox52.Text = "AVI" And ComboBox40.Text = "HUM" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "HUM" And TextBox52.Text = "PEM" Or TextBox52.Text = "HUM" And ComboBox40.Text = "PEM" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "HUM" And TextBox52.Text = "PEP" Or TextBox52.Text = "HUM" And ComboBox40.Text = "PEP" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "HUM" And TextBox52.Text = "PES" Or TextBox52.Text = "HUM" And ComboBox40.Text = "PES" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox40.Text = "HUM" And TextBox52.Text = "EAT" Or TextBox52.Text = "HUM" And ComboBox40.Text = "EAT" Then
            MsgBox("Несовместимые грузы в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        Mwt51 = CSng(TextBox45.Text)
        Mwt52 = CSng(TextBox40.Text)
        Mwt53 = CSng(TextBox50.Text)
        V51 = CSng(TextBox43.Text)
        V52 = CSng(TextBox39.Text)
        V53 = CSng(TextBox51.Text)
        If Mwt51 > 374 Then
            MsgBox("Превышение максимального веса в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If Mwt52 > 353 Then
            MsgBox("Превышение максимального веса в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If Mwt53 > 762 Then
            MsgBox("Превышение максимального веса в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V51 > 1.46 Then
            MsgBox("Превышение максимального объёма в позиции 51!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V52 > 1.38 Then
            MsgBox("Превышение максимального объёма в позиции 52!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If V53 > 2.88 Then
            MsgBox("Превышение максимального объёма в позиции 53!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        wt1 = CSng(TextBox8.Text)
        wt2 = CSng(TextBox12.Text)
        wt3 = CSng(TextBox16.Text)
        wt4 = CSng(TextBox26.Text)
        wt5 = CSng(TextBox21.Text)
        wt6 = CSng(TextBox36.Text)
        wt7 = CSng(TextBox30.Text)
        wt8 = CSng(TextBox45.Text)
        wt9 = CSng(TextBox40.Text)
        wt10 = CSng(TextBox50.Text)
        '11, 12 , 13
        If ComboBox1.Text = "F" And ComboBox5.Text = "F" And ComboBox9.Text = "F" Then
            Label45.Text = wt1 + wt2 + wt3
        ElseIf ComboBox1.Text = "F" And ComboBox5.Text = "F" Then
            Label45.Text = wt1 + wt2
        ElseIf ComboBox5.Text = "F" And ComboBox9.Text = "F" Then
            Label45.Text = wt2 + wt3
        ElseIf ComboBox1.Text = "F" And ComboBox9.Text = "F" Then
            Label45.Text = wt1 + wt3
        ElseIf ComboBox1.Text = "F" Then
            Label45.Text = wt1
        ElseIf ComboBox5.Text = "F" Then
            Label45.Text = wt2
        ElseIf ComboBox9.Text = "F" Then
            Label45.Text = wt3
        End If
        If ComboBox1.Text = "C" And ComboBox5.Text = "C" And ComboBox9.Text = "C" Then
            Label46.Text = wt1 + wt2 + wt3
        ElseIf ComboBox1.Text = "C" And ComboBox5.Text = "C" Then
            Label46.Text = wt1 + wt2
        ElseIf ComboBox5.Text = "C" And ComboBox9.Text = "C" Then
            Label46.Text = wt2 + wt3
        ElseIf ComboBox1.Text = "C" And ComboBox9.Text = "C" Then
            Label46.Text = wt1 + wt3
        ElseIf ComboBox1.Text = "C" Then
            Label46.Text = wt1
        ElseIf ComboBox5.Text = "C" Then
            Label46.Text = wt2
        ElseIf ComboBox9.Text = "C" Then
            Label46.Text = wt3
        End If
        If ComboBox1.Text = "M" And ComboBox5.Text = "M" And ComboBox9.Text = "M" Then
            Label48.Text = wt1 + wt2 + wt3
        ElseIf ComboBox1.Text = "M" And ComboBox5.Text = "M" Then
            Label48.Text = wt1 + wt2
        ElseIf ComboBox5.Text = "M" And ComboBox9.Text = "M" Then
            Label48.Text = wt2 + wt3
        ElseIf ComboBox1.Text = "M" And ComboBox9.Text = "M" Then
            Label48.Text = wt1 + wt3
        ElseIf ComboBox1.Text = "M" Then
            Label48.Text = wt1
        ElseIf ComboBox5.Text = "M" Then
            Label48.Text = wt2
        ElseIf ComboBox9.Text = "M" Then
            Label48.Text = wt3
        End If
        If ComboBox1.Text = "B" And ComboBox5.Text = "B" And ComboBox9.Text = "B" Then
            Label50.Text = wt1 + wt2 + wt3
        ElseIf ComboBox1.Text = "B" And ComboBox5.Text = "B" Then
            Label50.Text = wt1 + wt2
        ElseIf ComboBox5.Text = "B" And ComboBox9.Text = "B" Then
            Label50.Text = wt2 + wt3
        ElseIf ComboBox1.Text = "B" And ComboBox9.Text = "B" Then
            Label50.Text = wt1 + wt3
        ElseIf ComboBox1.Text = "B" Then
            Label50.Text = wt1
        ElseIf ComboBox5.Text = "B" Then
            Label50.Text = wt2
        ElseIf ComboBox9.Text = "B" Then
            Label50.Text = wt3
        End If
        ' 31 ,32
        If ComboBox18.Text = "F" And ComboBox16.Text = "F" Then
            Label58.Text = wt4 + wt5
        ElseIf ComboBox18.Text = "F" Then
            Label58.Text = wt4
        ElseIf ComboBox16.Text = "F" Then
            Label58.Text = wt5
        End If
        If ComboBox18.Text = "C" And ComboBox16.Text = "C" Then
            Label56.Text = wt4 + wt5
        ElseIf ComboBox18.Text = "C" Then
            Label56.Text = wt4
        ElseIf ComboBox16.Text = "C" Then
            Label56.Text = wt5
        End If
        If ComboBox18.Text = "M" And ComboBox16.Text = "M" Then
            Label54.Text = wt4 + wt5
        ElseIf ComboBox18.Text = "M" Then
            Label54.Text = wt4
        ElseIf ComboBox16.Text = "M" Then
            Label54.Text = wt5
        End If
        If ComboBox18.Text = "B" And ComboBox16.Text = "B" Then
            Label52.Text = wt4 + wt5
        ElseIf ComboBox18.Text = "B" Then
            Label52.Text = wt4
        ElseIf ComboBox16.Text = "B" Then
            Label52.Text = wt5
        End If
        '41,42
        If ComboBox26.Text = "F" And ComboBox24.Text = "F" Then
            Label66.Text = wt6 + wt7
        ElseIf ComboBox26.Text = "F" Then
            Label66.Text = wt6
        ElseIf ComboBox24.Text = "F" Then
            Label66.Text = wt7
        End If
        If ComboBox26.Text = "C" And ComboBox24.Text = "C" Then
            Label64.Text = wt6 + wt7
        ElseIf ComboBox26.Text = "C" Then
            Label64.Text = wt6
        ElseIf ComboBox24.Text = "C" Then
            Label64.Text = wt7
        End If
        If ComboBox26.Text = "M" And ComboBox24.Text = "M" Then
            Label62.Text = wt6 + wt7
        ElseIf ComboBox26.Text = "M" Then
            Label62.Text = wt6
        ElseIf ComboBox24.Text = "M" Then
            Label62.Text = wt7
        End If
        If ComboBox26.Text = "B" And ComboBox24.Text = "B" Then
            Label60.Text = wt6 + wt7
        ElseIf ComboBox26.Text = "B" Then
            Label60.Text = wt6
        ElseIf ComboBox24.Text = "B" Then
            Label60.Text = wt7
        End If
        'BULK
        If ComboBox34.Text = "F" And ComboBox32.Text = "F" And ComboBox37.Text = "F" Then
            Label85.Text = wt8 + wt9 + wt10
        ElseIf ComboBox34.Text = "F" And ComboBox32.Text = "F" Then
            Label85.Text = wt8 + wt9
        ElseIf ComboBox32.Text = "F" And ComboBox37.Text = "F" Then
            Label85.Text = wt9 + wt10
        ElseIf ComboBox34.Text = "F" And ComboBox37.Text = "F" Then
            Label85.Text = wt8 + wt10
        ElseIf ComboBox34.Text = "F" Then
            Label85.Text = wt8
        ElseIf ComboBox32.Text = "F" Then
            Label85.Text = wt9
        ElseIf ComboBox37.Text = "F" Then
            Label85.Text = wt10
        End If
        If ComboBox34.Text = "C" And ComboBox32.Text = "C" And ComboBox37.Text = "C" Then
            Label83.Text = wt8 + wt9 + wt10
        ElseIf ComboBox34.Text = "C" And ComboBox32.Text = "C" Then
            Label83.Text = wt8 + wt9
        ElseIf ComboBox32.Text = "C" And ComboBox37.Text = "C" Then
            Label83.Text = wt9 + wt10
        ElseIf ComboBox34.Text = "C" And ComboBox37.Text = "C" Then
            Label83.Text = wt8 + wt10
        ElseIf ComboBox34.Text = "C" Then
            Label83.Text = wt8
        ElseIf ComboBox32.Text = "C" Then
            Label83.Text = wt9
        ElseIf ComboBox37.Text = "C" Then
            Label83.Text = wt10
        End If
        If ComboBox34.Text = "M" And ComboBox32.Text = "M" And ComboBox37.Text = "M" Then
            Label81.Text = wt8 + wt9 + wt10
        ElseIf ComboBox34.Text = "M" And ComboBox32.Text = "M" Then
            Label81.Text = wt8 + wt9
        ElseIf ComboBox32.Text = "M" And ComboBox37.Text = "M" Then
            Label81.Text = wt9 + wt10
        ElseIf ComboBox34.Text = "M" And ComboBox37.Text = "M" Then
            Label81.Text = wt8 + wt10
        ElseIf ComboBox34.Text = "M" Then
            Label81.Text = wt8
        ElseIf ComboBox32.Text = "M" Then
            Label81.Text = wt9
        ElseIf ComboBox37.Text = "M" Then
            Label81.Text = wt10
        End If
        If ComboBox34.Text = "M" And ComboBox32.Text = "M" And ComboBox37.Text = "M" Then
            Label79.Text = wt8 + wt9 + wt10
        ElseIf ComboBox34.Text = "M" And ComboBox32.Text = "M" Then
            Label79.Text = wt8 + wt9
        ElseIf ComboBox32.Text = "M" And ComboBox37.Text = "M" Then
            Label79.Text = wt9 + wt10
        ElseIf ComboBox34.Text = "M" And ComboBox37.Text = "M" Then
            Label79.Text = wt8 + wt10
        ElseIf ComboBox34.Text = "M" Then
            Label79.Text = wt8
        ElseIf ComboBox32.Text = "M" Then
            Label79.Text = wt9
        ElseIf ComboBox37.Text = "M" Then
            Label79.Text = wt10
        End If
        a = CSng(Label45.Text)
        b = CSng(Label58.Text)
        c = CSng(Label66.Text)
        d = CSng(Label85.Text)
        Label69.Text = a + b + c + d
        k = CSng(Label46.Text)
        z = CSng(Label56.Text)
        f = CSng(Label64.Text)
        l = CSng(Label83.Text)
        Label71.Text = k + z + f + l
        g = CSng(Label48.Text)
        h = CSng(Label54.Text)
        i = CSng(Label62.Text)
        q = CSng(Label81.Text)
        Label73.Text = g + h + i + q
        t = CSng(Label50.Text)
        r = CSng(Label52.Text)
        w = CSng(Label60.Text)
        y = CSng(Label79.Text)
        Label75.Text = t + r + w + y
        TextBox6.Text = wt1 + wt2 + wt3
        TextBox27.Text = wt4 + wt5
        TextBox37.Text = wt6 + wt7
        TextBox46.Text = wt8 + wt9 + wt10
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox6.Text = "0"
        TextBox27.Text = "0"
        TextBox37.Text = "0"
        TextBox46.Text = "0"
        Label75.Text = "0"
        Label73.Text = "0"
        Label71.Text = "0"
        Label69.Text = "0"
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
        Label85.Text = "0"
        Label83.Text = "0"
        Label81.Text = "0"
        Label79.Text = "0"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox33.Text += 1
        Dim connection As New SqlConnection("Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True")
        Dim command As New SqlCommand("Insert into [Test].[dbo].[A320 lir]([flight_id1],[flight_route],[flight_bort],[config_id],[date_flight],[time_flight],[CPT1FWD],[dest11],[LIC11] ,[type11],[wt11],[stat11],[code11],[V11],[info11],[dest12],[LIC12],[type12],[wt12],[stat12],[code12],[V12],[info12],[dest13],[LIC13],[type13],[wt13],[stat13],[code13],[V13],[info13],[CPT3AFT],[dest31],[LIC31],[type31],[wt31],[stat31],[code31],[V31],[info31],[dest32],[LIC32],[type32],[wt32],[stat32],[code32],[V32],[info32],[CPT4AFT],[dest41],[LIC41],[type41],[wt41],[stat41],[code41],[V41],[info41],[dest42],[LIC42],[type42],[wt42],[stat42],[code42],[V42],[info42],[BULK],[dest51],[LIC51],[type51],[wt51],[stat51],[code51],[V51],[info51],[dest52],[LIC52],[type52],[wt52],[stat52],[code52],[V52],[info52],[dest53],[LIC53],[type53],[wt53],[stat53],[code53],[V53],[info53]) Values(@flight_id1, @flight_route, @flight_bort, @config_id, @date_flight, @time_flight, @CPT1FWD, @dest11, @LIC11, @type11, @wt11, @stat11, @code11, @V11, @info11, @dest12, @LIC12, @type12, @wt12, @stat12, @code12, @V12, @info12, @dest13, @LIC13, @type13, @wt13, @stat13, @code13, @V13, @info13, @CPT3AFT, @dest31, @LIC31, @type31, @wt31, @stat31, @code31, @V31, @info31, @dest32, @LIC32 , @type32 , @wt32 , @stat32 , @code32 , @V32, @info32, @CPT4AFT, @dest41, @LIC41 , @type41 , @wt41 , @stat41 , @code41 , @V41 , @info41 , @dest42 , @LIC42 , @type42 , @wt42 , @stat42, @code42 , @V42 , @info42, @BULK, @dest51, @LIC51, @type51, @wt51, @stat51, @code51, @V51, @info51, @dest52, @LIC52, @type52, @wt52, @stat52, @code52, @V52, @info52, @dest53, @LIC53, @type53, @wt53, @stat53, @code53, @V53, @info53)", connection)
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
        command.Parameters.AddWithValue("@dest13", TextBox15.Text)
        command.Parameters.AddWithValue("@LIC13", ComboBox9.Text)
        command.Parameters.AddWithValue("@type13", ComboBox10.Text)
        command.Parameters.AddWithValue("@wt13", TextBox16.Text)
        command.Parameters.AddWithValue("@stat13", ComboBox11.Text)
        command.Parameters.AddWithValue("@code13", ComboBox12.Text)
        command.Parameters.AddWithValue("@V13", TextBox17.Text)
        command.Parameters.AddWithValue("@info13", TextBox18.Text)
        command.Parameters.AddWithValue("@CPT3AFT", TextBox27.Text)
        command.Parameters.AddWithValue("@dest31", TextBox25.Text)
        command.Parameters.AddWithValue("@LIC31", ComboBox18.Text)
        command.Parameters.AddWithValue("@type31", ComboBox19.Text)
        command.Parameters.AddWithValue("@wt31", TextBox26.Text)
        command.Parameters.AddWithValue("@stat31", ComboBox20.Text)
        command.Parameters.AddWithValue("@code31", ComboBox17.Text)
        command.Parameters.AddWithValue("@V31", TextBox24.Text)
        command.Parameters.AddWithValue("@info31", TextBox23.Text)
        command.Parameters.AddWithValue("@dest32", TextBox22.Text)
        command.Parameters.AddWithValue("@LIC32", ComboBox16.Text)
        command.Parameters.AddWithValue("@type32", ComboBox15.Text)
        command.Parameters.AddWithValue("@wt32", TextBox21.Text)
        command.Parameters.AddWithValue("@stat32", ComboBox14.Text)
        command.Parameters.AddWithValue("@code32", ComboBox13.Text)
        command.Parameters.AddWithValue("@V32", TextBox20.Text)
        command.Parameters.AddWithValue("@info32", TextBox19.Text)
        command.Parameters.AddWithValue("@CPT4AFT", TextBox37.Text)
        command.Parameters.AddWithValue("@dest41", TextBox35.Text)
        command.Parameters.AddWithValue("@LIC41", ComboBox26.Text)
        command.Parameters.AddWithValue("@type41", ComboBox27.Text)
        command.Parameters.AddWithValue("@wt41", TextBox36.Text)
        command.Parameters.AddWithValue("@stat41", ComboBox28.Text)
        command.Parameters.AddWithValue("@code41", ComboBox25.Text)
        command.Parameters.AddWithValue("@V41", TextBox34.Text)
        command.Parameters.AddWithValue("@info41", TextBox32.Text)
        command.Parameters.AddWithValue("@dest42", TextBox31.Text)
        command.Parameters.AddWithValue("@LIC42", ComboBox24.Text)
        command.Parameters.AddWithValue("@type42", ComboBox23.Text)
        command.Parameters.AddWithValue("@wt42", TextBox30.Text)
        command.Parameters.AddWithValue("@stat42", ComboBox22.Text)
        command.Parameters.AddWithValue("@code42", ComboBox21.Text)
        command.Parameters.AddWithValue("@V42", TextBox29.Text)
        command.Parameters.AddWithValue("@info42", TextBox28.Text)
        command.Parameters.AddWithValue("@BULK", TextBox46.Text)
        command.Parameters.AddWithValue("@dest51", TextBox44.Text)
        command.Parameters.AddWithValue("@LIC51", ComboBox34.Text)
        command.Parameters.AddWithValue("@type51", ComboBox35.Text)
        command.Parameters.AddWithValue("@wt51", TextBox45.Text)
        command.Parameters.AddWithValue("@stat51", ComboBox36.Text)
        command.Parameters.AddWithValue("@code51", ComboBox33.Text)
        command.Parameters.AddWithValue("@V51", TextBox43.Text)
        command.Parameters.AddWithValue("@info51", TextBox42.Text)
        command.Parameters.AddWithValue("@dest52", TextBox41.Text)
        command.Parameters.AddWithValue("@LIC52", ComboBox32.Text)
        command.Parameters.AddWithValue("@type52", ComboBox31.Text)
        command.Parameters.AddWithValue("@wt52", TextBox40.Text)
        command.Parameters.AddWithValue("@stat52", ComboBox30.Text)
        command.Parameters.AddWithValue("@code52", ComboBox29.Text)
        command.Parameters.AddWithValue("@V52", TextBox39.Text)
        command.Parameters.AddWithValue("@info52", TextBox38.Text)
        command.Parameters.AddWithValue("@dest53", TextBox47.Text)
        command.Parameters.AddWithValue("@LIC53", ComboBox37.Text)
        command.Parameters.AddWithValue("@type53", ComboBox38.Text)
        command.Parameters.AddWithValue("@wt53", TextBox50.Text)
        command.Parameters.AddWithValue("@stat53", ComboBox39.Text)
        command.Parameters.AddWithValue("@code53", ComboBox40.Text)
        command.Parameters.AddWithValue("@V53", TextBox51.Text)
        command.Parameters.AddWithValue("@info53", TextBox52.Text)
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
        RichTextBox1.AppendText("CPT1FWD" + vbTab + "MAX" + "3402" + vbNewLine)
        RichTextBox1.AppendText(":" + "11P" + vbTab + vbTab + ComboBox2.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox3.Text + ":" + vbTab + TextBox7.Text + vbTab + ComboBox1.Text + "/" + TextBox8.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox4.Text + vbTab + TextBox9.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox8.Text + vbNewLine)
        RichTextBox1.AppendText("......................................................................................................" + vbNewLine)
        RichTextBox1.AppendText(":" + "12P" + vbTab + vbTab + ComboBox6.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox7.Text + ":" + vbTab + TextBox11.Text + vbTab + ComboBox5.Text + "/" + TextBox12.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox8.Text + vbTab + TextBox14.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox12.Text + vbNewLine)
        RichTextBox1.AppendText("......................................................................................................" + vbNewLine)
        RichTextBox1.AppendText(":" + "13P" + vbTab + vbTab + ComboBox10.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox11.Text + ":" + vbTab + TextBox15.Text + vbTab + ComboBox9.Text + "/" + TextBox16.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox12.Text + vbTab + TextBox18.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox16.Text + vbTab + vbTab + vbTab + "CPT1FWD" + vbTab + "TTL" + vbTab + TextBox6.Text + vbNewLine)
        RichTextBox1.AppendText("******************************************************************************************************************************" + vbNewLine)
        RichTextBox1.AppendText("CPT3AFT" + vbTab + vbTab + "MAX" + "2426" + vbNewLine)
        RichTextBox1.AppendText(":" + "31P" + vbTab + vbTab + ComboBox19.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox20.Text + ":" + vbTab + TextBox25.Text + vbTab + ComboBox18.Text + "/" + TextBox26.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox17.Text + vbTab + TextBox23.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox26.Text + vbNewLine)
        RichTextBox1.AppendText("......................................................................................................" + vbNewLine)
        RichTextBox1.AppendText(":" + "32P" + vbTab + vbTab + ComboBox15.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox14.Text + ":" + vbTab + TextBox22.Text + vbTab + ComboBox16.Text + "/" + TextBox21.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox13.Text + vbTab + TextBox19.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox21.Text + vbTab + vbTab + vbTab + "CPT3AFT" + vbTab + vbTab + "TTL" + vbTab + TextBox27.Text + vbNewLine)
        RichTextBox1.AppendText("******************************************************************************************************************************" + vbNewLine)
        RichTextBox1.AppendText("CPT4AFT" + vbTab + vbTab + "MAX" + "2110" + vbNewLine)
        RichTextBox1.AppendText(":" + "41P" + vbTab + vbTab + ComboBox27.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox28.Text + ":" + vbTab + TextBox35.Text + vbTab + ComboBox26.Text + "/" + TextBox36.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox25.Text + vbTab + TextBox32.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox36.Text + vbNewLine)
        RichTextBox1.AppendText("......................................................................................................" + vbNewLine)
        RichTextBox1.AppendText(":" + "42P" + vbTab + vbTab + ComboBox23.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox22.Text + ":" + vbTab + TextBox31.Text + vbTab + ComboBox24.Text + "/" + TextBox30.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox21.Text + vbTab + TextBox28.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox30.Text + vbTab + vbTab + vbTab + "CPT4AFT" + vbTab + vbTab + "TTL" + vbTab + TextBox37.Text + vbNewLine)
        RichTextBox1.AppendText("******************************************************************************************************************************" + vbNewLine)
        RichTextBox1.AppendText("BULK" + vbTab + "MAX" + "1489" + vbNewLine)
        RichTextBox1.AppendText(":" + "51P" + vbTab + vbTab + ComboBox35.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox36.Text + ":" + vbTab + TextBox44.Text + vbTab + ComboBox34.Text + "/" + TextBox45.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox33.Text + vbTab + TextBox42.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox45.Text + vbNewLine)
        RichTextBox1.AppendText("......................................................................................................" + vbNewLine)
        RichTextBox1.AppendText(":" + "52P" + vbTab + vbTab + ComboBox31.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox30.Text + ":" + vbTab + TextBox41.Text + vbTab + ComboBox32.Text + "/" + TextBox40.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox29.Text + vbTab + TextBox38.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox12.Text + vbNewLine)
        RichTextBox1.AppendText("......................................................................................................" + vbNewLine)
        RichTextBox1.AppendText(":" + "53P" + vbTab + vbTab + ComboBox38.Text + vbNewLine)
        RichTextBox1.AppendText(":" + ComboBox39.Text + ":" + vbTab + TextBox47.Text + vbTab + ComboBox37.Text + "/" + TextBox50.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "SPECS" + ":" + vbTab + ComboBox40.Text + vbTab + TextBox52.Text + vbNewLine)
        RichTextBox1.AppendText(":" + "REPORT" + ":" + vbTab + TextBox50.Text + vbTab + vbTab + vbTab + "BULK" + vbTab + "TTL" + vbTab + TextBox46.Text + vbNewLine)
        RichTextBox1.AppendText("******************************************************************************************************************************" + vbNewLine)
        Form1.TextBox36.Text = TextBox6.Text
        Form1.TextBox37.Text = TextBox27.Text
        Form1.TextBox38.Text = TextBox37.Text
        Form1.TextBox16.Text = TextBox46.Text
        Form1.TextBox6.Text = wt1 + wt2 + wt3 + wt4 + wt5 + wt6 + wt7 + wt8 + wt9 + wt10
    End Sub
End Class