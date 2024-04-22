Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form13
    Private dbconnections As New DatabaseConnections()
    Private connectionstr As String = dbconnections.GetConnectionString("stringconect_main")

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionstr)
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

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            CheckBox2.Checked = False
        ElseIf CheckBox2.Checked = True Then
            CheckBox1.Checked = False
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            CheckBox1.Checked = False
        ElseIf CheckBox1.Checked = True Then
            CheckBox2.Checked = False
        End If
    End Sub

End Class