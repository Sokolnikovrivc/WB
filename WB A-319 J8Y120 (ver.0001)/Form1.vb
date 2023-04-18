Imports System.Data.SqlClient
Imports System.IO.File
Public Class Form1
    Public C As Single
    Public Maxcomm As Single
    Public DOW As Single
    Public ZFW As Single
    Public TOF As Single
    Public TOW As Single
    Public TIF As Single
    Public LW As Single
    Public MZFW As Single
    Public MTOW As Single
    Public MLW As Single
    Public OA As Single
    Public OB As Single
    Public OC As Single
    Public OD As Single
    Public CPT1FWD As Single
    Public CPT4AFT As Single
    Public CPT5AFT As Single
    Public LIZFW As Single
    Public DOI As Single
    Public MACZFW As Single
    Public LITOW As Single
    Public LITOF As Single
    Public MACTOW As Single
    Public LILW As Single
    Public LITIF As Single
    Public MACLW As Single
    Public PAXwt As Single
    Public OA2 As Single
    Public OB2 As Single
    Public OC2 As Single
    Public OD2 As Single
    Public TTL As Single
    Public FWD1 As Single
    Public AFT2 As Single
    Public AFT4 As Single
    Public CPT6AFT As Single
    Public CPT7AFT As Single
    Public AFT6 As Single
    Public AFT7 As Single
    Public AFT8 As Single
    Public AFT9 As Single
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim connection As New SqlConnection("Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True")
            Dim command As New SqlCommand("select * from A319 where flight_id = @flight_id", connection)
            command.Parameters.Add("@flight_id", SqlDbType.VarChar).Value = TextBox5.Text
            Dim adapter As New SqlDataAdapter(command)
            Dim table As New DataTable
            adapter.Fill(table)
            TextBox4.Text = table.Rows(0)(0).ToString()
            TextBox2.Text = table.Rows(0)(1).ToString()
            TextBox3.Text = table.Rows(0)(8).ToString()
            TextBox48.Text = table.Rows(0)(5).ToString()
            TextBox49.Text = table.Rows(0)(6).ToString()
            TextBox13.Text = table.Rows(0)(2).ToString()
            TextBox19.Text = table.Rows(0)(3).ToString()
            TextBox18.Text = table.Rows(0)(4).ToString()
            ComboBox1.Focus()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim filename As String = Application.StartupPath & "\test.log"
        Dim sw As IO.StreamWriter = AppendText(filename)
        sw.WriteLine(Now() & "" & "sample log file entry")
        sw.Close()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TTL = CSng(TextBox6.Text)
        PAXwt = CSng(TextBox7.Text)
        TextBox8.Text = TTL + PAXwt
        C = CSng(TextBox8.Text)
        Maxcomm = CSng(TextBox9.Text)
        If Maxcomm < C Then
            MsgBox("Фактическая коммерческая загрузка превышает предельную!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox6.Focus()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        C = TextBox8.Text
        DOW = TextBox11.Text
        TextBox12.Text = C + DOW
        ZFW = CSng(TextBox12.Text)
        MZFW = CSng(TextBox13.Text)
        If MZFW < ZFW Then
            MsgBox("Фактическое значение ZFW превышает предельное MZFW!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox12.Text = ""
            TextBox12.Focus()
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox12.Text = ZFW
        TOF = CSng(TextBox15.Text)
        TextBox17.Text = ZFW + TOF
        TOW = CSng(TextBox17.Text)
        MTOW = CSng(TextBox19.Text)
        If MTOW < TOW Then
            MsgBox("Фактическое значение TOW превышает предельное MTOW!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox17.Text = ""
            TextBox17.Focus()
        End If
        Label29.Text = MTOW - TOW
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox17.Text = TOW
        TIF = CSng(TextBox20.Text)
        TextBox21.Text = TOW - TIF
        LW = CSng(TextBox21.Text)
        MLW = CSng(TextBox18.Text)
        If MLW < LW Then
            MsgBox("Фактическое значение TLW превышает предельное LW!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox21.Text = ""
            TextBox21.Focus()
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim connection As New SqlConnection("Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True")
        Dim command As New SqlCommand("select * from [A319 Wt] where Flight_bort = @Flight_bort and crew = @crew", connection)
        command.Parameters.Add("@Flight_bort", SqlDbType.VarChar).Value = TextBox4.Text
        command.Parameters.Add("@crew", SqlDbType.VarChar).Value = ComboBox1.Text
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable
        adapter.Fill(table)
        TextBox23.Text = table.Rows(0)(2).ToString()
        TextBox14.Text = table.Rows(0)(3).ToString()
        TextBox24.Text = table.Rows(0)(4).ToString()
        TextBox25.Text = table.Rows(0)(5).ToString()
        TextBox11.Text = TextBox14.Text
        DOW = CSng(TextBox14.Text)
        MZFW = CSng(TextBox13.Text)
        DOI = CSng(TextBox24.Text)
        TextBox9.Text = MZFW - DOW
        TextBox6.Focus()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim MAXwtOA As Integer = registryKey.GetValue("TextBox8", "")
        Dim OA1 As Single = registryKey.GetValue("TextBox12", "")
        registryKey.Close()
        OA2 = CSng(TextBox28.Text)
        TextBox32.Text = OA2 * OA1
        OA = CSng(TextBox32.Text)
        If OA2 > MAXwtOA Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox28.Text = ""
            TextBox32.Text = ""
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim MAXwtOB As Integer = registryKey.GetValue("TextBox9", "")
        Dim OB1 As Single = registryKey.GetValue("TextBox13", "")
        registryKey.Close()
        OB2 = CSng(TextBox29.Text)
        TextBox33.Text = OB2 * OB1
        OB = CSng(TextBox33.Text)
        If OB2 > MAXwtOB Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox29.Text = ""
            TextBox33.Text = ""
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim MAXwtOC As Integer = registryKey.GetValue("TextBox10", "")
        Dim OC1 As Single = registryKey.GetValue("TextBox14", "")
        registryKey.Close()
        OC2 = CSng(TextBox30.Text)
        TextBox34.Text = OC2 * OC1
        OC = CSng(TextBox34.Text)
        If OC2 > MAXwtOC Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox30.Text = ""
            TextBox34.Text = ""
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim MAXwtOD As Integer = registryKey.GetValue("TextBox11", "")
        Dim OD1 As Single = registryKey.GetValue("TextBox15", "")
        registryKey.Close()
        OD2 = CSng(TextBox31.Text)
        PAXwt = CSng(TextBox7.Text)
        TextBox35.Text = OD2 * OD1
        OD = CSng(TextBox35.Text)
        If OD2 > MAXwtOD Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox31.Text = ""
            TextBox35.Text = ""
        End If
        If PAXwt < OA2 + OB2 + OC2 + OD2 Then
            MsgBox("Превышение значения введённого веса пассажиров!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox31.Text = ""
            TextBox30.Text = ""
            TextBox29.Text = ""
            TextBox28.Text = ""
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim FWD As Single = registryKey.GetValue("TextBox36", "")
        Dim FWDkey As Single = registryKey.GetValue("TextBox16", "")
        registryKey.Close()
        FWD1 = CSng(TextBox36.Text)
        TextBox39.Text = FWD1 * FWD
        CPT1FWD = CSng(TextBox39.Text)
        If FWD1 > FWDkey Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox36.Text = ""
            TextBox39.Text = ""
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim AFT1 As Single = registryKey.GetValue("TextBox35", "")
        Dim AFT1key As Single = registryKey.GetValue("TextBox17", "")
        registryKey.Close()
        AFT2 = CSng(TextBox37.Text)
        TextBox40.Text = AFT1 * AFT2
        CPT4AFT = CSng(TextBox40.Text)
        If AFT2 > AFT1key Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox37.Text = ""
            TextBox40.Text = ""
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim AFT3 As Single = registryKey.GetValue("TextBox34", "")
        Dim AFT3key As Single = registryKey.GetValue("TextBox18", "")
        registryKey.Close()
        AFT4 = CSng(TextBox38.Text)
        TTL = CSng(TextBox6.Text)
        TextBox41.Text = AFT3 * AFT4
        CPT5AFT = CSng(TextBox40.Text)
        If AFT4 > AFT3key Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox38.Text = ""
            TextBox41.Text = ""
        End If
        If TextBox1.Text = "A-319" Then
            If FWD1 + AFT2 + AFT4 + AFT7 + AFT9 > TTL Then
                MsgBox("Превышение значения введённого веса грузов!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
                TextBox36.Text = ""
                TextBox37.Text = ""
                TextBox38.Text = ""
            End If
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        LIZFW = CSng(TextBox42.Text)
        TextBox42.Text = DOI + OA + OB + OC + OD + CPT1FWD + CPT4AFT + CPT5AFT
    End Sub
    Public Function Preobraz_index()
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim RefSTA As Integer = registryKey.GetValue("TextBox29", "")
        Dim K As Integer = registryKey.GetValue("TextBox30", "")
        Dim C As Integer = registryKey.GetValue("TextBox31", "")
        Dim CAX As Single = registryKey.GetValue("TextBox32", "")
        Dim LEMAC As Single = registryKey.GetValue("TextBox33", "")
        LIZFW = CSng(TextBox42.Text)
        ZFW = CSng(TextBox12.Text)
        MACTOW = CSng(TextBox45.Text)
        LITOW = CSng(TextBox44.Text)
        TOW = CSng(TextBox17.Text)
        LILW = CSng(TextBox46.Text)
        LITIF = CSng(TextBox27.Text)
        registryKey.Close()
        Dim Answer As Single = ((((C * (LIZFW - K)) / ZFW) + RefSTA - LEMAC) / CAX) * 100
        Return Answer
    End Function
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        TextBox43.Text = Preobraz_index()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        LITOW = CSng(TextBox44.Text)
        LITOF = CSng(TextBox26.Text)
        TextBox44.Text = DOI + OA + OB + OC + OD + CPT1FWD + CPT4AFT + CPT5AFT + LITOF
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim RefSTA As Integer = registryKey.GetValue("TextBox29", "")
        Dim K As Integer = registryKey.GetValue("TextBox30", "")
        Dim C As Integer = registryKey.GetValue("TextBox31", "")
        Dim CAX As Single = registryKey.GetValue("TextBox32", "")
        Dim LEMAC As Single = registryKey.GetValue("TextBox33", "")
        registryKey.Close()
        MACTOW = CSng(TextBox45.Text)
        LITOW = CSng(TextBox44.Text)
        TOW = CSng(TextBox17.Text)
        TextBox45.Text = ((((C * (LITOW - K)) / TOW) + RefSTA - LEMAC) / CAX) * 100
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        LILW = CSng(TextBox46.Text)
        LITIF = CSng(TextBox27.Text)
        TextBox46.Text = DOI + OA + OB + OC + OD + CPT1FWD + CPT4AFT + CPT5AFT + LITOF + LITIF
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim RefSTA As Integer = registryKey.GetValue("TextBox29", "")
        Dim K As Integer = registryKey.GetValue("TextBox30", "")
        Dim C As Integer = registryKey.GetValue("TextBox31", "")
        Dim CAX As Single = registryKey.GetValue("TextBox32", "")
        Dim LEMAC As Single = registryKey.GetValue("TextBox33", "")
        registryKey.Close()
        MACLW = CSng(TextBox47.Text)
        LILW = CSng(TextBox46.Text)
        LW = CSng(TextBox21.Text)
        TextBox47.Text = ((((C * (LILW - K)) / LW) + RefSTA - LEMAC) / CAX) * 100
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim connection As New SqlConnection("Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True")
        Dim command As New SqlCommand("Insert into [Test].[dbo].[A319 Act] ([flight1_Id] ,[flight_Route] ,[flight1_bort] ,[config_id] ,[date_flight] ,[time_flight] ,[ttlKZ],[ZFW],[TOW],[TIF],[LW],[LIZFW],[LITOW],[LILW],[MACZFW],[MACTOW],[MACLW],[TTL_load],[PAX_wt],[DOW],[code_crew],[code_kitchen],[DOI],[MAC],[TOF],[LITOF],[LITIF],[OA],[OB],[OC],[OD],[CPT1FWD],[CPT4AFT],[CPT5AFT] ) Values(@flight1_Id, @flight_Route, @flight1_bort, @config_id, @date_flight, @time_flight, @ttlKZ, @ZFW, @TOW, @TIF, @LW, @LIZFW, @LITOW, @LILW, @MACZFW, @MACTOW, @MACLW, @TTL_load, @PAX_wt, @DOW, @code_crew, @code_kitchen, @DOI, @MAC, @TOF, @LITOF, @LITIF, @OA, @OB, @OC, @OD, @CPT1FWD, @CPT4AFT, @CPT5AFT)", connection)
        command.Parameters.AddWithValue("@flight1_Id", TextBox5.Text)
        command.Parameters.AddWithValue("@flight_Route", TextBox3.Text)
        command.Parameters.AddWithValue("@flight1_bort", TextBox4.Text)
        command.Parameters.AddWithValue("@config_id", TextBox2.Text)
        command.Parameters.AddWithValue("@date_flight", TextBox48.Text)
        command.Parameters.AddWithValue("@time_flight", TextBox49.Text)
        command.Parameters.AddWithValue("@ttlKZ", TextBox8.Text)
        command.Parameters.AddWithValue("@ZFW", TextBox12.Text)
        command.Parameters.AddWithValue("@TOW", TextBox17.Text)
        command.Parameters.AddWithValue("@TIF", TextBox20.Text)
        command.Parameters.AddWithValue("@LW", TextBox21.Text)
        command.Parameters.AddWithValue("@LIZFW", TextBox42.Text)
        command.Parameters.AddWithValue("@LITOW", TextBox44.Text)
        command.Parameters.AddWithValue("@LILW", TextBox46.Text)
        command.Parameters.AddWithValue("@MACZFW", TextBox43.Text)
        command.Parameters.AddWithValue("@MACTOW", TextBox45.Text)
        command.Parameters.AddWithValue("@MACLW", TextBox47.Text)
        command.Parameters.AddWithValue("@TTL_load", TextBox6.Text)
        command.Parameters.AddWithValue("@PAX_wt", TextBox7.Text)
        command.Parameters.AddWithValue("@DOW", TextBox14.Text)
        command.Parameters.AddWithValue("@code_crew", ComboBox1.Text)
        command.Parameters.AddWithValue("@code_kitchen", TextBox23.Text)
        command.Parameters.AddWithValue("@DOI", TextBox24.Text)
        command.Parameters.AddWithValue("@MAC", TextBox25.Text)
        command.Parameters.AddWithValue("@TOF", TextBox15.Text)
        command.Parameters.AddWithValue("@LITOF", TextBox26.Text)
        command.Parameters.AddWithValue("@LITIF", TextBox27.Text)
        command.Parameters.AddWithValue("@OA", TextBox28.Text)
        command.Parameters.AddWithValue("@OB", TextBox29.Text)
        command.Parameters.AddWithValue("@OC", TextBox30.Text)
        command.Parameters.AddWithValue("@OD", TextBox31.Text)
        command.Parameters.AddWithValue("@CPT1FWD", TextBox36.Text)
        command.Parameters.AddWithValue("@CPT4AFT", TextBox37.Text)
        command.Parameters.AddWithValue("@CPT5AFT", TextBox38.Text)
        connection.Open()
        command.ExecuteNonQuery()
        connection.Close()
        MessageBox.Show("Запись в Базу Данных")

    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Form3.Show()
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Form2.Show()
        Form2.TextBox5.Text = TextBox5.Text
        Form2.TextBox3.Text = TextBox3.Text
        Form2.TextBox48.Text = TextBox48.Text
        Form2.TextBox49.Text = TextBox49.Text
        Form2.TextBox2.Text = TextBox2.Text
        Form2.TextBox4.Text = TextBox4.Text
        Form2.TextBox7.Text = Mid(TextBox3.Text, 5, 3)
        Form2.TextBox11.Text = Mid(TextBox3.Text, 5, 3)
        Form2.TextBox22.Text = Mid(TextBox3.Text, 5, 3)
        Form2.TextBox19.Text = Mid(TextBox3.Text, 5, 3)
        Form2.TextBox31.Text = Mid(TextBox3.Text, 5, 3)
        Form2.TextBox28.Text = Mid(TextBox3.Text, 5, 3)
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Form5.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form6.Show()
        Form6.TextBox5.Text = TextBox5.Text
        Form6.TextBox3.Text = TextBox3.Text
        Form6.TextBox48.Text = TextBox48.Text
        Form6.TextBox49.Text = TextBox49.Text
        Form6.TextBox2.Text = TextBox2.Text
        Form6.TextBox4.Text = TextBox4.Text
        Form6.Label17.Text = MTOW - TOW
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            Microsoft.Win32.Registry.CurrentUser.DeleteSubKey("FormDatachanges")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Const AFT6 As Single = 0.00447
        AFT7 = CSng(TextBox16.Text)
        TextBox10.Text = AFT6 * AFT7
        CPT6AFT = CSng(TextBox10.Text)
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Const AFT8 As Single = 0.00447
        AFT9 = CSng(TextBox50.Text)
        TextBox22.Text = AFT8 * AFT9
        CPT7AFT = CSng(TextBox22.Text)
    End Sub
End Class


