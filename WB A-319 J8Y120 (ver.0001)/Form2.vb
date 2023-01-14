
Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Mwt11 As Single
        Dim Mwt12 As Single
        Dim V11 As Single
        Dim V12 As Single
        If ComboBox4.Text = "RFL" And ComboBox8.Text = "ROX" Or ComboBox4.Text = "ROX" And ComboBox8.Text = "RFL" Then
            MsgBox("Не совместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RSC" And ComboBox8.Text = "ROX" Or ComboBox4.Text = "ROX" And ComboBox8.Text = "RSC" Then
            MsgBox("Не совместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RFW" And ComboBox8.Text = "RCM" Or ComboBox4.Text = "RCM" And ComboBox8.Text = "RFW" Then
            MsgBox("Не совместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And ComboBox8.Text = "RPB" Or ComboBox4.Text = "RPB" And ComboBox8.Text = "AVI" Then
            MsgBox("Не совместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And ComboBox8.Text = "RIS" Or ComboBox4.Text = "RIS" And ComboBox8.Text = "AVI" Then
            MsgBox("Не совместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And ComboBox8.Text = "ICE" Or ComboBox4.Text = "ICE" And ComboBox8.Text = "AVI" Then
            MsgBox("Не совместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RPB" And ComboBox8.Text = "HEG" Or ComboBox4.Text = "HEG" And ComboBox8.Text = "RPB" Then
            MsgBox("Не совместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RIS" And ComboBox8.Text = "HEG" Or ComboBox4.Text = "HEG" And ComboBox8.Text = "RIS" Then
            MsgBox("Не совместимые грузы в позициях 11 и 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RPB" And TextBox9.Text = "PEM" Or TextBox9.Text = "RPB" And ComboBox4.Text = "PEM" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RPB" And TextBox9.Text = "PEP" Or TextBox9.Text = "RPB" And ComboBox4.Text = "PEP" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RPB" And TextBox9.Text = "PES" Or TextBox9.Text = "RPB" And ComboBox4.Text = "PES" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RPB" And TextBox9.Text = "EAT" Or TextBox9.Text = "RPB" And ComboBox4.Text = "EAT" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RIS" And TextBox9.Text = "PEM" Or TextBox9.Text = "RIS" And ComboBox4.Text = "PEM" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RIS" And TextBox9.Text = "PEP" Or TextBox9.Text = "RIS" And ComboBox4.Text = "PEP" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RIS" And TextBox9.Text = "PES" Or TextBox9.Text = "RIS" And ComboBox4.Text = "PES" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "RIS" And TextBox9.Text = "EAT" Or TextBox9.Text = "RIS" And ComboBox4.Text = "EAT" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And TextBox9.Text = "AVI" Or TextBox9.Text = "AVI" And ComboBox4.Text = "AVI" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "AVI" And TextBox9.Text = "HUM" Or TextBox9.Text = "AVI" And ComboBox4.Text = "HUM" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "HUM" And TextBox9.Text = "PEM" Or TextBox9.Text = "HUM" And ComboBox4.Text = "PEM" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "HUM" And TextBox9.Text = "PEP" Or TextBox9.Text = "HUM" And ComboBox4.Text = "PEP" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "HUM" And TextBox9.Text = "PES" Or TextBox9.Text = "HUM" And ComboBox4.Text = "PES" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox4.Text = "HUM" And TextBox9.Text = "EAT" Or TextBox9.Text = "HUM" And ComboBox4.Text = "EAT" Then
            MsgBox("Не совместимые грузы в позициии 11!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RPB" And TextBox14.Text = "PEM" Or TextBox14.Text = "RPB" And ComboBox8.Text = "PEM" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RPB" And TextBox14.Text = "PEP" Or TextBox14.Text = "RPB" And ComboBox8.Text = "PEP" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RPB" And TextBox14.Text = "PES" Or TextBox14.Text = "RPB" And ComboBox8.Text = "PES" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RPB" And TextBox14.Text = "EAT" Or TextBox14.Text = "RPB" And ComboBox8.Text = "EAT" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RIS" And TextBox14.Text = "PEM" Or TextBox14.Text = "RIS" And ComboBox8.Text = "PEM" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RIS" And TextBox14.Text = "PEP" Or TextBox14.Text = "RIS" And ComboBox8.Text = "PEP" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RIS" And TextBox14.Text = "PES" Or TextBox14.Text = "RIS" And ComboBox8.Text = "PES" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "RIS" And TextBox14.Text = "EAT" Or TextBox14.Text = "RIS" And ComboBox8.Text = "EAT" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "AVI" And TextBox14.Text = "AVI" Or TextBox14.Text = "AVI" And ComboBox8.Text = "AVI" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "AVI" And TextBox14.Text = "HUM" Or TextBox14.Text = "AVI" And ComboBox8.Text = "HUM" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "HUM" And TextBox14.Text = "PEM" Or TextBox14.Text = "HUM" And ComboBox8.Text = "PEM" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "HUM" And TextBox14.Text = "PEP" Or TextBox14.Text = "HUM" And ComboBox8.Text = "PEP" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "HUM" And TextBox14.Text = "PES" Or TextBox14.Text = "HUM" And ComboBox8.Text = "PES" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        If ComboBox8.Text = "HUM" And TextBox14.Text = "EAT" Or TextBox14.Text = "HUM" And ComboBox8.Text = "EAT" Then
            MsgBox("Не совместимые грузы в позициии 12!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
        End If
        Mwt11 = CSng(TextBox8.Text)
        Mwt12 = CSng(TextBox12.Text)
        V11 = CSng(TextBox10.Text)
        V12 = CSng(TextBox13.Text)
        If Mwt11 > 1045 Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox8.Text = ""
        End If
        If Mwt12 > 1223 Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox12.Text = ""
        End If
        If V11 > 4.09 Then
            MsgBox("Превышение максимального объёма!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox10.Text = ""
        End If
        If V12 > 4.42 Then
            MsgBox("Превышение максимального объёма!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox13.Text = ""
        End If
    End Sub
End Class