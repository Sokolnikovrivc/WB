Imports System.Timers
Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports log4net
Imports System.Data.Common


Public Class Form12
    Private timer As Timer
    Private dbconnections As New DatabaseConnections()
    Private connectionstr As String = dbconnections.GetConnectionString("stringconect_main")
    Private Sub Form12_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Инициализация таймера
        timer = New Timer()
        timer.Interval = 1000 ' Интервал обновления в миллисекундах
        AddHandler timer.Elapsed, AddressOf TimerElapsed
        timer.Start()
        Try
            Using connection As New SqlConnection(connectionstr)
                Dim Command As New SqlCommand("GetDataA319", connection)
                Command.CommandType = CommandType.StoredProcedure
                Dim adapter As New SqlDataAdapter(Command)
                Dim table As New DataTable
                adapter.Fill(table)
                DataGridView1.DataSource = table
            End Using

        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub TimerElapsed(sender As Object, e As ElapsedEventArgs)
        ' Метод, вызываемый при срабатывании таймера
        UpdateDateTime()
    End Sub
    Public Sub disp_data1()
        Using connection As New SqlConnection(connectionstr)
            Dim command As New SqlCommand("GetDataMethodView", connection)
            Dim adapter3 As New SqlDataAdapter(command)
            Dim table As New DataTable()
            adapter3.Fill(table)
            DataGridView1.DataSource = table
        End Using
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 AndAlso e.RowIndex < DataGridView1.Rows.Count Then
            Dim selectedRow As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            Try
                TextBox1.Text = selectedRow.Cells(0).Value.ToString()
                TextBox2.Text = selectedRow.Cells(1).Value.ToString()
                TextBox3.Text = selectedRow.Cells(2).Value.ToString()
                TextBox4.Text = selectedRow.Cells(3).Value.ToString()
                TextBox5.Text = selectedRow.Cells(4).Value.ToString()
                TextBox6.Text = selectedRow.Cells(5).Value.ToString()
                TextBox7.Text = selectedRow.Cells(6).Value.ToString()
            Catch ex As Exception
                MsgBox("Error: " & ex.ToString())
            End Try
        Else
            ' Индекс строки вышел за пределы допустимого диапазона
            ' Вы можете добавить обработку ошибки или вывести сообщение об ошибке
        End If
    End Sub

    Private Sub UpdateDateTime()
        ' Получаем текущую дату и время
        Dim currentDate As DateTime = DateTime.Now

        ' Преобразуем дату и время в строку
        Dim currentDateTimeString As String = currentDate.ToString()

        ' Обновляем значение времени на форме
        UpdateLabel(currentDateTimeString)
    End Sub

    Private Sub UpdateLabel(dateTimeString As String)
        ' Обновление значения времени на форме через делегат
        Try
            If InvokeRequired Then
                Invoke(Sub() UpdateLabel(dateTimeString))
            Else
                Label1.Text = dateTimeString
            End If
        Catch ex As Exception
            MsgBox(ex)
        End Try

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form13.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Using connection As New SqlConnection(connectionstr)
                Dim command As New SqlCommand("UpdateDataF12", connection)
                command.CommandType = CommandType.StoredProcedure
                command.Parameters.AddWithValue("@flight_id", TextBox1.Text)
                command.Parameters.AddWithValue("@flight_route", TextBox2.Text)
                command.Parameters.AddWithValue("@date_flight", TextBox3.Text)
                command.Parameters.AddWithValue("@time_flight", TextBox4.Text)
                command.Parameters.AddWithValue("@type_Aircraft", TextBox5.Text)
                command.Parameters.AddWithValue("@flight_bort", TextBox6.Text)
                command.Parameters.AddWithValue("@config_id", TextBox7.Text)
                'command.ExecuteNonQuery()
                Dim dt As New DataTable
                Dim adapter As New SqlDataAdapter(command)
                adapter.Fill(dt)

                MessageBox.Show("Обновление данных")
                disp_data1()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Using connection As New SqlConnection(connectionstr)
                connection.Open()
                Dim command As New SqlCommand("DELETE FROM [Test].[dbo].[A319] WHERE flight_id = @flight_id", connection)
                command.Parameters.AddWithValue("@flight_id", TextBox1.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Удаление данных")
                disp_data1()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox5.Text = "A-319" Then
            Form10.TextBox1.Text = TextBox5.Text
            Form10.TextBox2.Text = TextBox7.Text
            Form10.TextBox3.Text = TextBox6.Text
            Form10.Show()
        ElseIf TextBox5.Text = "A-320" Then
            Form11.TextBox1.Text = TextBox5.Text
            Form11.TextBox2.Text = TextBox7.Text
            Form11.TextBox3.Text = TextBox6.Text
            Form11.Show()
        End If
    End Sub
End Class