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
    Public BULK As Single
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
    Public connectionString As String = "Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Label36.Text = registryKey.GetValue("Label27", "")
        Label37.Text = registryKey.GetValue("Label26", "")
        Label38.Text = registryKey.GetValue("Label25", "")
        Label23.Text = registryKey.GetValue("Label40", "")
        registryKey.Close()
        Dim filename As String = Application.StartupPath & "\test.log"
        Dim sw As IO.StreamWriter = AppendText(filename)
        sw.WriteLine(Now() & "" & "sample log file entry")
        sw.Close()
        ComboBox1.Items.Clear()
        combo() ' заполняем ComboBox данными из базы данных
    End Sub
    Public Sub combo()
        If TextBox1.Text = "A-319" Then
            Try
                Using connection As SqlConnection = New SqlConnection(connectionString)
                    connection.Open()
                    Dim command As New SqlCommand("SELECT * FROM [Test].[dbo].[A319 Wt] where config_id = @config_id", connection)
                    command.Parameters.Add("@config_id", SqlDbType.VarChar).Value = TextBox2.Text
                    Dim reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        ComboBox1.Items.Add(reader("crew").ToString())
                    End While
                End Using
            Catch ex As Exception
                MsgBox("Error: " & ex.ToString())
            End Try
        ElseIf TextBox1.Text = "A-320" Then ' добавляем проверку на второй вариант
            Try
                Using connection As SqlConnection = New SqlConnection(connectionString)
                    connection.Open()
                    Dim command As New SqlCommand("SELECT * FROM [Test].[dbo].[A320 Wt] where config_id = @config_id", connection)
                    command.Parameters.Add("@config_id", SqlDbType.VarChar).Value = TextBox2.Text
                    Dim reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        ComboBox1.Items.Add(reader("crew").ToString())
                    End While
                End Using
            Catch ex As Exception
                MsgBox("Error: " & ex.ToString())
            End Try
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "A-320" Then
            Label23.Visible = True
            TextBox16.Visible = True
            Button25.Visible = True
            TextBox10.Visible = True
        End If
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
        If TextBox1.Text = "A-319" Then
            Try
                Using connection As SqlConnection = New SqlConnection(connectionString)
                    connection.Open()
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
                End Using
            Catch ex As Exception
                MsgBox("Error: " & ex.ToString())
            End Try
        ElseIf TextBox1.Text = "A-320" Then
            Try
                Using connection As SqlConnection = New SqlConnection(connectionString)
                    connection.Open()
                    Dim command As New SqlCommand("select * from [A320 Wt] where Flight_bort = @Flight_bort and crew = @crew", connection)
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
                End Using
            Catch ex As Exception
                MsgBox("Error: " & ex.ToString())
            End Try
        End If
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
        DOI = CSng(TextBox24.Text)
        OA = CSng(TextBox32.Text)
        OB = CSng(TextBox33.Text)
        OC = CSng(TextBox34.Text)
        OD = CSng(TextBox35.Text)
        CPT1FWD = CSng(TextBox39.Text)
        CPT4AFT = CSng(TextBox40.Text)
        CPT5AFT = CSng(TextBox41.Text)
        BULK = CSng(TextBox10.Text)
        TextBox42.Text = CStr(DOI + OA + OB + OC + OD + CPT1FWD + CPT4AFT + CPT5AFT + BULK)
        LIZFW = CSng(TextBox42.Text)
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
        TextBox44.Text = DOI + OA + OB + OC + OD + CPT1FWD + CPT4AFT + CPT5AFT + BULK + LITOF
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
        TextBox46.Text = DOI + OA + OB + OC + OD + CPT1FWD + CPT4AFT + CPT5AFT + BULK + LITOF + LITIF
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
        Dim command As New SqlCommand("Insert into [Test].[dbo].[A319 Act] ([flight1_Id] ,[flight_Route] ,[flight1_bort] ,[config_id] ,[date_flight] ,[time_flight] ,[ttlKZ],[ZFW],[TOW],[TIF],[LW],[LIZFW],[LITOW],[LILW],[MACZFW],[MACTOW],[MACLW],[TTL_load],[PAX_wt],[DOW],[code_crew],[code_kitchen],[DOI],[MAC],[TOF],[LITOF],[LITIF],[OA],[OB],[OC],[OD],[CPT1FWD],[CPT4AFT],[CPT5AFT],[BLK],[LICPT1FWD],[LICPT4AFT],[LICPT5AFT],[LIBLK],[LIOA],[LIOB],[LIOC],[LIOD],[MAXKZ],[MAXZFW],[MAXTOW],[MAXLW],[UNDERLOAD],[C],[Y],[CARGO],[M],[B],[VZR],[RB],[RM],[TTL],[type_Aircraft]) Values(@flight1_Id, @flight_Route, @flight1_bort, @config_id, @date_flight, @time_flight, @ttlKZ, @ZFW, @TOW, @TIF, @LW, @LIZFW, @LITOW, @LILW, @MACZFW, @MACTOW, @MACLW, @TTL_load, @PAX_wt, @DOW, @code_crew, @code_kitchen, @DOI, @MAC, @TOF, @LITOF, @LITIF, @OA, @OB, @OC, @OD, @CPT1FWD, @CPT4AFT, @CPT5AFT, @BLK, @LICPT1FWD, @LICPT4AFT, @LICPT5AFT, @LIBLK, @LIOA, @LIOB, @LIOC, @LIOD, @MAXKZ, @MAXZFW, @MAXTOW, @MAXLW, @UNDERLOAD, @C, @Y, @CARGO, @M, @B, @VZR, @RB, @RM, @TTL, @type_Aircraft)", connection)
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
        command.Parameters.AddWithValue("@BLK", TextBox16.Text)
        command.Parameters.AddWithValue("@LICPT1FWD", TextBox39.Text)
        command.Parameters.AddWithValue("@LICPT4AFT", TextBox40.Text)
        command.Parameters.AddWithValue("@LICPT5AFT", TextBox41.Text)
        command.Parameters.AddWithValue("@LIBLK", TextBox10.Text)
        command.Parameters.AddWithValue("@LIOA", TextBox32.Text)
        command.Parameters.AddWithValue("@LIOB", TextBox33.Text)
        command.Parameters.AddWithValue("@LIOC", TextBox34.Text)
        command.Parameters.AddWithValue("@LIOD", TextBox35.Text)
        command.Parameters.AddWithValue("@MAXKZ", TextBox9.Text)
        command.Parameters.AddWithValue("@MAXZFW", TextBox13.Text)
        command.Parameters.AddWithValue("@MAXTOW", TextBox19.Text)
        command.Parameters.AddWithValue("@MAXLW", TextBox18.Text)
        command.Parameters.AddWithValue("@UNDERLOAD", Label29.Text)
        command.Parameters.AddWithValue("@C", Label57.Text)
        command.Parameters.AddWithValue("@Y", Label56.Text)
        command.Parameters.AddWithValue("@CARGO", Label71.Text)
        command.Parameters.AddWithValue("@M", Label73.Text)
        command.Parameters.AddWithValue("@B", Label75.Text)
        command.Parameters.AddWithValue("@VZR", Label50.Text)
        command.Parameters.AddWithValue("@RB", Label51.Text)
        command.Parameters.AddWithValue("@RM", Label52.Text)
        command.Parameters.AddWithValue("@TTL", Label53.Text)
        command.Parameters.AddWithValue("@type_Aircraft", TextBox1.Text)
        connection.Open()
        command.ExecuteNonQuery()
        connection.Close()
        MessageBox.Show("Запись в Базу Данных")
    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Form3.Show()
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        If TextBox1.Text = "A-320" Then
            Try
                Using connection As SqlConnection = New SqlConnection(connectionString)
                    connection.Open()
                    Dim command As New SqlCommand("SELECT TOP 1 * FROM [Test].[dbo].[A320 lir] where flight_id1 = @flight_id1 and flight_bort = @flight_bort ORDER BY CreatedAt DESC", connection)
                    command.Parameters.Add("@flight_id1", SqlDbType.VarChar).Value = TextBox5.Text
                    command.Parameters.Add("@flight_bort", SqlDbType.VarChar).Value = TextBox4.Text
                    Dim adapter As New SqlDataAdapter(command)
                    Dim table As New DataTable
                    adapter.Fill(table)
                    Form9.TextBox5.Text = table.Rows(0)(0).ToString()
                    Form9.TextBox3.Text = table.Rows(0)(1).ToString()
                    Form9.TextBox4.Text = table.Rows(0)(2).ToString()
                    Form9.TextBox2.Text = table.Rows(0)(3).ToString()
                    Form9.TextBox48.Text = table.Rows(0)(4).ToString()
                    Form9.TextBox49.Text = table.Rows(0)(5).ToString()
                    Form9.TextBox6.Text = table.Rows(0)(6).ToString()
                    Form9.TextBox7.Text = table.Rows(0)(7).ToString()
                    Form9.ComboBox1.Text = table.Rows(0)(8).ToString()
                    Form9.ComboBox2.Text = table.Rows(0)(9).ToString()
                    Form9.TextBox8.Text = table.Rows(0)(10).ToString()
                    Form9.ComboBox3.Text = table.Rows(0)(11).ToString()
                    Form9.ComboBox4.Text = table.Rows(0)(12).ToString()
                    Form9.TextBox10.Text = table.Rows(0)(13).ToString()
                    Form9.TextBox9.Text = table.Rows(0)(14).ToString()
                    Form9.TextBox11.Text = table.Rows(0)(15).ToString()
                    Form9.ComboBox5.Text = table.Rows(0)(16).ToString()
                    Form9.ComboBox6.Text = table.Rows(0)(17).ToString()
                    Form9.TextBox12.Text = table.Rows(0)(18).ToString()
                    Form9.ComboBox7.Text = table.Rows(0)(19).ToString()
                    Form9.ComboBox8.Text = table.Rows(0)(20).ToString()
                    Form9.TextBox13.Text = table.Rows(0)(21).ToString()
                    Form9.TextBox14.Text = table.Rows(0)(22).ToString()
                    Form9.TextBox15.Text = table.Rows(0)(23).ToString()
                    Form9.ComboBox9.Text = table.Rows(0)(24).ToString()
                    Form9.ComboBox10.Text = table.Rows(0)(25).ToString()
                    Form9.TextBox16.Text = table.Rows(0)(26).ToString()
                    Form9.ComboBox11.Text = table.Rows(0)(27).ToString()
                    Form9.ComboBox12.Text = table.Rows(0)(28).ToString()
                    Form9.TextBox17.Text = table.Rows(0)(29).ToString()
                    Form9.TextBox18.Text = table.Rows(0)(30).ToString()
                    Form9.TextBox27.Text = table.Rows(0)(31).ToString()
                    Form9.TextBox25.Text = table.Rows(0)(32).ToString()
                    Form9.ComboBox18.Text = table.Rows(0)(33).ToString()
                    Form9.ComboBox19.Text = table.Rows(0)(34).ToString()
                    Form9.TextBox26.Text = table.Rows(0)(35).ToString()
                    Form9.ComboBox20.Text = table.Rows(0)(36).ToString()
                    Form9.ComboBox17.Text = table.Rows(0)(37).ToString()
                    Form9.TextBox24.Text = table.Rows(0)(38).ToString()
                    Form9.TextBox23.Text = table.Rows(0)(39).ToString()
                    Form9.TextBox22.Text = table.Rows(0)(40).ToString()
                    Form9.ComboBox16.Text = table.Rows(0)(41).ToString()
                    Form9.ComboBox15.Text = table.Rows(0)(42).ToString()
                    Form9.TextBox21.Text = table.Rows(0)(43).ToString()
                    Form9.ComboBox14.Text = table.Rows(0)(44).ToString()
                    Form9.ComboBox13.Text = table.Rows(0)(45).ToString()
                    Form9.TextBox20.Text = table.Rows(0)(46).ToString()
                    Form9.TextBox19.Text = table.Rows(0)(47).ToString()
                    Form9.TextBox37.Text = table.Rows(0)(48).ToString()
                    Form9.TextBox35.Text = table.Rows(0)(49).ToString()
                    Form9.ComboBox26.Text = table.Rows(0)(50).ToString()
                    Form9.ComboBox27.Text = table.Rows(0)(51).ToString()
                    Form9.TextBox36.Text = table.Rows(0)(52).ToString()
                    Form9.ComboBox28.Text = table.Rows(0)(53).ToString()
                    Form9.ComboBox25.Text = table.Rows(0)(54).ToString()
                    Form9.TextBox34.Text = table.Rows(0)(55).ToString()
                    Form9.TextBox32.Text = table.Rows(0)(56).ToString()
                    Form9.TextBox31.Text = table.Rows(0)(57).ToString()
                    Form9.ComboBox24.Text = table.Rows(0)(58).ToString()
                    Form9.ComboBox23.Text = table.Rows(0)(59).ToString()
                    Form9.TextBox30.Text = table.Rows(0)(60).ToString()
                    Form9.ComboBox22.Text = table.Rows(0)(61).ToString()
                    Form9.ComboBox21.Text = table.Rows(0)(62).ToString()
                    Form9.TextBox29.Text = table.Rows(0)(63).ToString()
                    Form9.TextBox28.Text = table.Rows(0)(64).ToString()
                    Form9.TextBox46.Text = table.Rows(0)(65).ToString()
                    Form9.TextBox44.Text = table.Rows(0)(66).ToString()
                    Form9.ComboBox34.Text = table.Rows(0)(67).ToString()
                    Form9.ComboBox35.Text = table.Rows(0)(68).ToString()
                    Form9.TextBox45.Text = table.Rows(0)(69).ToString()
                    Form9.ComboBox36.Text = table.Rows(0)(70).ToString()
                    Form9.ComboBox33.Text = table.Rows(0)(71).ToString()
                    Form9.TextBox43.Text = table.Rows(0)(72).ToString()
                    Form9.TextBox42.Text = table.Rows(0)(73).ToString()
                    Form9.TextBox41.Text = table.Rows(0)(74).ToString()
                    Form9.ComboBox32.Text = table.Rows(0)(75).ToString()
                    Form9.ComboBox31.Text = table.Rows(0)(76).ToString()
                    Form9.TextBox40.Text = table.Rows(0)(77).ToString()
                    Form9.ComboBox30.Text = table.Rows(0)(78).ToString()
                    Form9.ComboBox29.Text = table.Rows(0)(79).ToString()
                    Form9.TextBox39.Text = table.Rows(0)(80).ToString()
                    Form9.TextBox38.Text = table.Rows(0)(81).ToString()
                    Form9.TextBox47.Text = table.Rows(0)(81).ToString()
                    Form9.ComboBox37.Text = table.Rows(0)(83).ToString()
                    Form9.ComboBox38.Text = table.Rows(0)(84).ToString()
                    Form9.TextBox50.Text = table.Rows(0)(85).ToString()
                    Form9.ComboBox39.Text = table.Rows(0)(86).ToString()
                    Form9.ComboBox40.Text = table.Rows(0)(87).ToString()
                    Form9.TextBox51.Text = table.Rows(0)(88).ToString()
                    Form9.TextBox52.Text = table.Rows(0)(89).ToString()
                End Using
            Catch ex As Exception
                MsgBox("По данному рейсу записей не найдено!", vbInformation, "Информация")
            End Try
            Form9.Show()
            Form9.TextBox5.Text = TextBox5.Text
            Form9.TextBox3.Text = TextBox3.Text
            Form9.TextBox48.Text = TextBox48.Text
            Form9.TextBox49.Text = TextBox49.Text
            Form9.TextBox2.Text = TextBox2.Text
            Form9.TextBox4.Text = TextBox4.Text
            Form9.TextBox1.Text = TextBox1.Text
            Form9.TextBox7.Text = Mid(TextBox3.Text, 5, 3)
            Form9.TextBox11.Text = Mid(TextBox3.Text, 5, 3)
            Form9.TextBox15.Text = Mid(TextBox3.Text, 5, 3)
            Form9.TextBox25.Text = Mid(TextBox3.Text, 5, 3)
            Form9.TextBox22.Text = Mid(TextBox3.Text, 5, 3)
            Form9.TextBox35.Text = Mid(TextBox3.Text, 5, 3)
            Form9.TextBox31.Text = Mid(TextBox3.Text, 5, 3)
            Form9.TextBox53.Text = Mid(TextBox3.Text, 5, 3)
            Form9.TextBox41.Text = Mid(TextBox3.Text, 5, 3)
            Form9.TextBox47.Text = Mid(TextBox3.Text, 5, 3)
        End If
        If TextBox1.Text = "A-319" Then
            Try
                Using connection As SqlConnection = New SqlConnection(connectionString)
                    connection.Open()
                    Dim command As New SqlCommand("SELECT TOP 1 * FROM [Test].[dbo].[A319 lir] where flight_id1 = @flight_id1 and flight_bort = @flight_bort ORDER BY CreatedAt DESC", connection)
                    command.Parameters.Add("@flight_id1", SqlDbType.VarChar).Value = TextBox5.Text
                    command.Parameters.Add("@flight_bort", SqlDbType.VarChar).Value = TextBox4.Text
                    Dim adapter As New SqlDataAdapter(command)
                    Dim table As New DataTable
                    adapter.Fill(table)
                    Form2.TextBox5.Text = table.Rows(0)(0).ToString()
                    Form2.TextBox3.Text = table.Rows(0)(1).ToString()
                    Form2.TextBox4.Text = table.Rows(0)(2).ToString()
                    Form2.TextBox2.Text = table.Rows(0)(3).ToString()
                    Form2.TextBox48.Text = table.Rows(0)(4).ToString()
                    Form2.TextBox49.Text = table.Rows(0)(5).ToString()
                    Form2.TextBox6.Text = table.Rows(0)(6).ToString()
                    Form2.TextBox7.Text = table.Rows(0)(7).ToString()
                    Form2.ComboBox1.Text = table.Rows(0)(8).ToString()
                    Form2.ComboBox2.Text = table.Rows(0)(9).ToString()
                    Form2.TextBox8.Text = table.Rows(0)(10).ToString()
                    Form2.ComboBox3.Text = table.Rows(0)(11).ToString()
                    Form2.ComboBox4.Text = table.Rows(0)(12).ToString()
                    Form2.TextBox10.Text = table.Rows(0)(13).ToString()
                    Form2.TextBox9.Text = table.Rows(0)(14).ToString()
                    Form2.TextBox11.Text = table.Rows(0)(15).ToString()
                    Form2.ComboBox5.Text = table.Rows(0)(16).ToString()
                    Form2.ComboBox6.Text = table.Rows(0)(17).ToString()
                    Form2.TextBox12.Text = table.Rows(0)(18).ToString()
                    Form2.ComboBox7.Text = table.Rows(0)(19).ToString()
                    Form2.ComboBox8.Text = table.Rows(0)(20).ToString()
                    Form2.TextBox13.Text = table.Rows(0)(21).ToString()
                    Form2.TextBox14.Text = table.Rows(0)(22).ToString()
                    Form2.TextBox15.Text = table.Rows(0)(23).ToString()
                    Form2.TextBox22.Text = table.Rows(0)(24).ToString()
                    Form2.ComboBox14.Text = table.Rows(0)(25).ToString()
                    Form2.ComboBox15.Text = table.Rows(0)(26).ToString()
                    Form2.TextBox23.Text = table.Rows(0)(27).ToString()
                    Form2.ComboBox16.Text = table.Rows(0)(28).ToString()
                    Form2.ComboBox13.Text = table.Rows(0)(29).ToString()
                    Form2.TextBox21.Text = table.Rows(0)(30).ToString()
                    Form2.TextBox20.Text = table.Rows(0)(31).ToString()
                    Form2.TextBox19.Text = table.Rows(0)(32).ToString()
                    Form2.ComboBox12.Text = table.Rows(0)(33).ToString()
                    Form2.ComboBox11.Text = table.Rows(0)(34).ToString()
                    Form2.TextBox18.Text = table.Rows(0)(35).ToString()
                    Form2.ComboBox10.Text = table.Rows(0)(36).ToString()
                    Form2.ComboBox9.Text = table.Rows(0)(37).ToString()
                    Form2.TextBox17.Text = table.Rows(0)(38).ToString()
                    Form2.TextBox16.Text = table.Rows(0)(39).ToString()
                    Form2.TextBox24.Text = table.Rows(0)(40).ToString()
                    Form2.TextBox31.Text = table.Rows(0)(41).ToString()
                    Form2.ComboBox22.Text = table.Rows(0)(42).ToString()
                    Form2.ComboBox23.Text = table.Rows(0)(43).ToString()
                    Form2.TextBox32.Text = table.Rows(0)(44).ToString()
                    Form2.ComboBox24.Text = table.Rows(0)(45).ToString()
                    Form2.ComboBox21.Text = table.Rows(0)(46).ToString()
                    Form2.TextBox30.Text = table.Rows(0)(47).ToString()
                    Form2.TextBox29.Text = table.Rows(0)(48).ToString()
                    Form2.TextBox28.Text = table.Rows(0)(49).ToString()
                    Form2.ComboBox20.Text = table.Rows(0)(50).ToString()
                    Form2.ComboBox19.Text = table.Rows(0)(51).ToString()
                    Form2.TextBox27.Text = table.Rows(0)(52).ToString()
                    Form2.ComboBox18.Text = table.Rows(0)(53).ToString()
                    Form2.ComboBox17.Text = table.Rows(0)(54).ToString()
                    Form2.TextBox26.Text = table.Rows(0)(55).ToString()
                    Form2.TextBox25.Text = table.Rows(0)(56).ToString()
                End Using
            Catch ex As Exception
                MsgBox("По данному рейсу записей не найдено!", vbInformation, "Информация")
            End Try
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
        End If
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Form5.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form6.TextBox1.Text = TextBox1.Text
        Form6.TextBox5.Text = TextBox5.Text
        Form6.TextBox3.Text = TextBox3.Text
        Form6.TextBox48.Text = TextBox48.Text
        Form6.TextBox49.Text = TextBox49.Text
        Form6.TextBox2.Text = TextBox2.Text
        Form6.TextBox4.Text = TextBox4.Text
        Form6.Label17.Text = Label29.Text
        Form6.Label20.Text = Label57.Text
        Form6.Label21.Text = Label56.Text
        Form6.Label71.Text = Label71.Text
        Form6.Label73.Text = Label73.Text
        Form6.Label75.Text = Label75.Text
        Form6.Show()
    End Sub
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            Microsoft.Win32.Registry.CurrentUser.DeleteSubKey("FormDatachanges")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        Dim BULK1 As Single = registryKey.GetValue("TextBox58", "")
        Dim BULKkey As Single = registryKey.GetValue("TextBox57", "")
        registryKey.Close()
        AFT7 = CSng(TextBox16.Text)
        TTL = CSng(TextBox6.Text)
        BULK = CSng(TextBox10.Text)
        TextBox10.Text = AFT7 * BULK1
        If AFT7 > BULKkey Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox10.Text = ""
            TextBox16.Text = ""
        End If
        If TextBox1.Text = "A-320" Then
            If FWD1 + AFT2 + AFT4 + AFT7 + AFT9 > TTL Then
                MsgBox("Превышение значения введённого веса грузов!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
                TextBox36.Text = ""
                TextBox37.Text = ""
                TextBox38.Text = ""
            End If
        End If
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Const AFT8 As Single = 0.00447
        AFT9 = CSng(TextBox50.Text)
        TextBox22.Text = AFT8 * AFT9
        CPT7AFT = CSng(TextBox22.Text)
    End Sub
End Class


