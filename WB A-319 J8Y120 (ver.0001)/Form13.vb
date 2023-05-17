Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form13
    Public connectionString As String = "Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True"

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("Insert into [Test].[dbo].[A319] ([flight_bort],[config_id],[MZFW],[MTOW],[MLW],[date_flight],[time_flight],[flight_id],[flight_route],[type_Aircraft]) Values(@flight_bort, @config_id, @MZFW, @MTOW, @MLW, @date_flight, @time_flight, @flight_id, @flight_route, @type_Aircraft)", connection)
                command.Parameters.AddWithValue("@flight_bort", TextBox3.Text)
                command.Parameters.AddWithValue("@config_id", TextBox6.Text)
                command.Parameters.AddWithValue("@MZFW", TextBox8.Text)
                command.Parameters.AddWithValue("@MTOW", TextBox9.Text)
                command.Parameters.AddWithValue("@MLW", TextBox7.Text)
                command.Parameters.AddWithValue("@date_flight", TextBox4.Text)
                command.Parameters.AddWithValue("@time_flight", TextBox5.Text)
                command.Parameters.AddWithValue("@flight_id", TextBox1.Text)
                command.Parameters.AddWithValue("@flight_route", TextBox2.Text)
                If CheckBox1.Checked Then
                    command.Parameters.AddWithValue("@type_Aircraft", Label9.Text)
                ElseIf CheckBox2.Checked Then
                    command.Parameters.AddWithValue("@type_Aircraft", Label10.Text)
                End If
                command.ExecuteNonQuery()
                MessageBox.Show("Запись в Базу Данных")
                Form12.disp_data1()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub
End Class