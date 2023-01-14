
Imports System.ComponentModel.Design
Imports System.Data.SqlClient
Imports System.Drawing.Printing
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim connection As New SqlConnection("Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True")
        Dim command As New SqlCommand("select * from A319 where flight_id = @flight_id", connection)
        command.Parameters.Add("@flight_id", SqlDbType.VarChar).Value = TextBox5.Text
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable
        adapter.Fill(table)
        TextBox3.Text = table.Rows(0)(1).ToString()
        TextBox4.Text = table.Rows(0)(2).ToString()
        TextBox2.Text = table.Rows(0)(3).ToString()
        TextBox13.Text = table.Rows(0)(4).ToString()
        TextBox19.Text = table.Rows(0)(5).ToString()
        TextBox18.Text = table.Rows(0)(6).ToString()
        TextBox48.Text = table.Rows(0)(7).ToString()
        TextBox49.Text = table.Rows(0)(8).ToString()
        TextBox22.Focus()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim filename As String = Application.StartupPath & "\test.log"
        Dim sw As IO.StreamWriter = AppendText(filename)
        sw.WriteLine(Now() & "" & "sample log file entry")
        sw.Close()
        'TODO: данная строка кода позволяет загрузить данные в таблицу "TestDataSet5.fuel_810". При необходимости она может быть перемещена или удалена.
        Me.Fuel_810TableAdapter.Fill(Me.TestDataSet5.fuel_810)
        'TODO: данная строка кода позволяет загрузить данные в таблицу "TestDataSet4.fuel_0800". При необходимости она может быть перемещена или удалена.
        Me.Fuel_0800TableAdapter.Fill(Me.TestDataSet4.fuel_0800)
        'TODO: данная строка кода позволяет загрузить данные в таблицу "TestDataSet3.fuel_079". При необходимости она может быть перемещена или удалена.
        Me.Fuel_079TableAdapter.Fill(Me.TestDataSet3.fuel_079)
        'TODO: данная строка кода позволяет загрузить данные в таблицу "TestDataSet2.Weight_pax". При необходимости она может быть перемещена или удалена.
        Me.Weight_paxTableAdapter.Fill(Me.TestDataSet2.Weight_pax)

    End Sub

    Private Sub FillByToolStripButton_Click(sender As Object, e As EventArgs)
        Try
            Me.Weight_paxTableAdapter.FillBy(Me.TestDataSet2.Weight_pax)
        Catch ex As System.Exception
            System.Windows.Forms.MessageBox.Show(ex.Message)
        End Try
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
        TOF = CSng(TextBox10.Text)
        TextBox10.Text = TextBox15.Text
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
        TIF = CSng(TextBox16.Text)
        TextBox16.Text = TextBox20.Text
        TextBox26.Focus()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim connection As New SqlConnection("Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True")
        Dim command As New SqlCommand("select * from [A319 Wt] where Flight_bort = @Flight_bort and crew = @crew", connection)
        command.Parameters.Add("@Flight_bort", SqlDbType.VarChar).Value = TextBox4.Text
        command.Parameters.Add("@crew", SqlDbType.VarChar).Value = TextBox22.Text
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
        Const OA1 As Single = -0.00705
        OA2 = CSng(TextBox28.Text)
        TextBox32.Text = OA2 * OA1
        OA = CSng(TextBox32.Text)
        If OA2 > 680 Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox28.Text = ""
            TextBox32.Text = ""
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Const OB1 As Single = -0.00325
        OB2 = CSng(TextBox29.Text)
        TextBox33.Text = OB2 * OB1
        OB = CSng(TextBox33.Text)
        If OB2 > 3570 Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox29.Text = ""
            TextBox33.Text = ""
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Const OC1 As Single = 0.00231
        OC2 = CSng(TextBox30.Text)
        TextBox34.Text = OC2 * OC1
        OC = CSng(TextBox34.Text)
        If OC2 > 3570 Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox30.Text = ""
            TextBox34.Text = ""
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Const OD1 As Single = 0.00726
        OD2 = CSng(TextBox31.Text)
        PAXwt = CSng(TextBox7.Text)
        TextBox35.Text = OD2 * OD1
        OD = CSng(TextBox35.Text)
        If OD2 > 3060 Then
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
        Const FWD As Single = -0.00563
        FWD1 = CSng(TextBox36.Text)
        TextBox39.Text = FWD1 * FWD
        CPT1FWD = CSng(TextBox39.Text)
        If FWD1 > 2268 Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox36.Text = ""
            TextBox39.Text = ""
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Const AFT1 As Single = 0.00447
        AFT2 = CSng(TextBox37.Text)
        TextBox40.Text = AFT1 * AFT2
        CPT4AFT = CSng(TextBox40.Text)
        If AFT2 > 3021 Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox37.Text = ""
            TextBox40.Text = ""
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Const AFT3 As Single = 0.0084
        AFT4 = CSng(TextBox38.Text)
        TTL = CSng(TextBox6.Text)
        TextBox41.Text = AFT3 * AFT4
        CPT5AFT = CSng(TextBox40.Text)
        If AFT4 > 1497 Then
            MsgBox("Превышение максимального веса!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox38.Text = ""
            TextBox41.Text = ""
        End If
        If FWD1 + AFT2 + AFT4 > TTL Then
            MsgBox("Превышение значения введённого веса грузов!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox36.Text = ""
            TextBox37.Text = ""
            TextBox38.Text = ""
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        LIZFW = CSng(TextBox42.Text)
        TextBox42.Text = DOI + OA + OB + OC + OD + CPT1FWD + CPT4AFT + CPT5AFT
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Const M As Single = 100000
        Const K As Single = 50
        Const RefSTA As Single = 1725
        Const LEMAC As Single = 1620.2
        Const CAX As Single = 419.4
        MACZFW = CSng(TextBox43.Text)
        LIZFW = CSng(TextBox42.Text)
        ZFW = CSng(TextBox12.Text)
        TextBox43.Text = ((((M * (LIZFW - K)) / ZFW) + RefSTA - LEMAC) / CAX) * 100
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        LITOW = CSng(TextBox44.Text)
        LITOF = CSng(TextBox26.Text)
        TextBox44.Text = DOI + OA + OB + OC + OD + CPT1FWD + CPT4AFT + CPT5AFT + LITOF
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Const M As Single = 100000
        Const K As Single = 50
        Const RefSTA As Single = 1725
        Const LEMAC As Single = 1620.2
        Const CAX As Single = 419.4
        MACTOW = CSng(TextBox45.Text)
        LITOW = CSng(TextBox44.Text)
        TOW = CSng(TextBox17.Text)
        TextBox45.Text = ((((M * (LITOW - K)) / TOW) + RefSTA - LEMAC) / CAX) * 100
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        LILW = CSng(TextBox46.Text)
        LITIF = CSng(TextBox27.Text)
        TextBox46.Text = DOI + OA + OB + OC + OD + CPT1FWD + CPT4AFT + CPT5AFT + LITOF + LITIF
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Const M As Single = 100000
        Const K As Single = 50
        Const RefSTA As Single = 1725
        Const LEMAC As Single = 1620.2
        Const CAX As Single = 419.4
        MACLW = CSng(TextBox47.Text)
        LILW = CSng(TextBox46.Text)
        LW = CSng(TextBox21.Text)
        TextBox47.Text = ((((M * (LILW - K)) / LW) + RefSTA - LEMAC) / CAX) * 100
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim connection As New SqlConnection("Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True")
        Dim command As New SqlCommand("Insert into [Test].[dbo].[A319 Act] ([flight1_Id] ,[flight_Route] ,[flight1_bort] ,[config_id] ,[date_flight] ,[time_flight] ,[ttlKZ],[ZFW],[TOW],[TIF],[LW],[LIZFW],[LITOW],[LILW],[MACZFW],[MACTOW],[MACLW]) Values(@flight1_Id, @flight_Route, @flight1_bort, @config_id, @date_flight, @time_flight, @ttlKZ, @ZFW, @TOW, @TIF, @LW, @LIZFW, @LITOW, @LILW, @MACZFW, @MACTOW, @MACLW)", connection)
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
        connection.Open()
        command.ExecuteNonQuery()
        connection.Close()
        MessageBox.Show("Запись в Базу Данных")
        RichTextBox1.Visible = True
        RichTextBox1.Clear()
        RichTextBox1.AppendText("LOADSHEET" + vbTab + vbTab + "CHECKED" + vbTab + vbTab + "APPROVED" + vbTab + vbTab + "EDNO" + vbNewLine)
        RichTextBox1.AppendText("ALL WEIGHTS IN KILOS" + vbNewLine)
        RichTextBox1.AppendText("FROM/TO" + vbTab + "FLIGHT" + vbTab + "A/C REG" + vbTab + vbTab + "VERSION" + vbTab + "CREW" + vbTab + "DATE" + vbTab + vbTab + "TIME" + vbNewLine)
        RichTextBox1.AppendText(TextBox3.Text + vbTab + TextBox5.Text + vbTab + TextBox4.Text + vbTab + vbTab + vbTab + TextBox2.Text + vbTab + vbTab + TextBox22.Text + vbTab + TextBox48.Text + vbTab + TextBox49.Text + vbNewLine)
        RichTextBox1.AppendText(vbTab + vbTab + vbTab + vbTab + "WEIGHT" + vbTab + "DISTRIBUTION" + vbNewLine)
        RichTextBox1.AppendText("LOAD IN COMPARTMENTS" + vbTab + TextBox6.Text + vbTab + vbTab + "1/" + TextBox36.Text + vbTab + "4/" + TextBox37.Text + vbTab + "5/" + TextBox38.Text + vbNewLine)
        RichTextBox1.AppendText("PASSENGER" + vbTab + vbTab + vbTab + TextBox7.Text + vbTab + vbTab + Label50.Text + "/" + Label51.Text + "/" + Label52.Text + vbTab + vbTab + vbTab + "TTL" + Label53.Text + vbNewLine)
        RichTextBox1.AppendText("*************************************************************************" + vbNewLine)
        RichTextBox1.AppendText("TOTAL TRAFFIC LOAD" + vbTab + vbTab + TextBox8.Text + vbNewLine)
        RichTextBox1.AppendText("DRY OPERATING WEIGHT" + vbTab + TextBox11.Text + vbNewLine)
        RichTextBox1.AppendText("ZERO FUEL WEIGHT" + vbTab + vbTab + TextBox12.Text + vbNewLine)
        RichTextBox1.AppendText("TAKE OFF FUEL" + vbTab + vbTab + vbTab + TextBox15.Text + vbNewLine)
        RichTextBox1.AppendText("TAKE OFF WEIGHT ACTUAL" + vbTab + TextBox17.Text + vbNewLine)
        RichTextBox1.AppendText("TRIP FUEL" + vbTab + vbTab + vbTab + TextBox20.Text + vbNewLine)
        RichTextBox1.AppendText("LANDING WEIGHT ACTUAL" + vbTab + TextBox21.Text + vbNewLine)
        RichTextBox1.AppendText("BALANCE AND SEATING CONDITIONS" + vbNewLine)
        RichTextBox1.AppendText("DOI" + vbTab + TextBox24.Text + vbTab + "LIZFW" + vbTab + TextBox42.Text + vbNewLine)
        RichTextBox1.AppendText("LITOW" + vbTab + TextBox44.Text + vbTab + "MACZFW" + vbTab + TextBox43.Text + vbNewLine)
        RichTextBox1.AppendText("MACTOW" + vbTab + vbTab + TextBox45.Text + vbNewLine)
        RichTextBox1.AppendText("0A/" + TextBox28.Text + vbTab + "0B/" + TextBox29.Text + vbTab + "0C/" + TextBox30.Text + vbTab + "0D/" + TextBox31.Text + vbTab + vbNewLine)
        RichTextBox1.AppendText("CABIN AREA TRIM" + vbNewLine)
        RichTextBox1.AppendText("UNDERLOAD BEFORE LMC" + vbNewLine)
        RichTextBox1.AppendText("*************************************************************************")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            Dim PrinDoc As New PrintDocument
            AddHandler PrinDoc.PrintPage, AddressOf Me.PrintText
            PrinDoc.Print()
            PrintPreviewDialog1.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Beep()
        End Try
    End Sub
    Private Sub PrintText(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        Try
            e.Graphics.DrawString(RichTextBox1.Text, New Font("TimesNewRoman", 12, FontStyle.Regular), Brushes.Black, New Point(0, 0))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Beep()
        End Try
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Try
            e.Graphics.DrawString(RichTextBox1.Text, New Font("TimesNewRoman", 12, FontStyle.Regular), Brushes.Black, New Point(0, 0))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Beep()
        End Try
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
        Form2.TextBox6.Text = TextBox36.Text
        Form2.TextBox15.Text = TextBox37.Text
        Form2.TextBox24.Text = TextBox38.Text
    End Sub
End Class


