
Imports System.ComponentModel.Design
Imports System.Data.SqlClient

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
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        Dim TTL, PAXWt As Single
        TTL = TextBox6.Text
        PAXWt = TextBox7.Text
        TextBox8.Text = TTL + PAXWt
        C = CSng(TextBox8.Text)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        C = TextBox8.Text
        DOW = TextBox11.Text
        TextBox12.Text = C + DOW
        ZFW = CSng(TextBox12.Text)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox12.Text = ZFW
        TOF = CSng(TextBox15.Text)
        TextBox17.Text = ZFW + TOF
        TOW = CSng(TextBox17.Text)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox17.Text = TOW
        TIF = CSng(TextBox20.Text)
        TextBox21.Text = TOW - TIF
        LW = CSng(TextBox21.Text)
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
        TextBox9.Text = MZFW - DOW
        Maxcomm = CSng(TextBox9.Text)
        If Maxcomm < C Then
            MsgBox("Фактическая коммерческая загрузка превышает предельную!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox6.Focus()
        End If
    End Sub
End Class

