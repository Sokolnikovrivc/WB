Imports System.ComponentModel.Design
Imports System.Data.SqlClient
Imports System.Windows.Forms
Public Class Form10
    Public connectionString As String = "Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True"
    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                ' Создать команду для выборки данных
                Dim command3 As New SqlCommand("SELECT FuelWeight AS Вес_топлива, CenterOfGravityIndex as Индекс FROM [Test].[dbo].[fuel_079]", connection)
                ' Создать адаптер данных и заполнить DataTable
                Dim adapter3 As New SqlDataAdapter(command3)
                Dim table3 As New DataTable
                adapter3.Fill(table3)
                ' Установить DataTable в качестве источника данных для DataGridView
                DataGridView2.DataSource = table3
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                ' Создать команду для выборки данных
                Dim command4 As New SqlCommand("SELECT FuelWeight AS Вес_топлива, CenterOfGravityIndex as Индекс FROM [Test].[dbo].[fuel_080]", connection)
                ' Создать адаптер данных и заполнить DataTable
                Dim adapter4 As New SqlDataAdapter(command4)
                Dim table4 As New DataTable
                adapter4.Fill(table4)
                ' Установить DataTable в качестве источника данных для DataGridView
                DataGridView3.DataSource = table4
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                ' Создать команду для выборки данных
                Dim command5 As New SqlCommand("SELECT FuelWeight AS Вес_топлива, CenterOfGravityIndex as Индекс FROM [Test].[dbo].[fuel_810]", connection)
                ' Создать адаптер данных и заполнить DataTable
                Dim adapter5 As New SqlDataAdapter(command5)
                Dim table5 As New DataTable
                adapter5.Fill(table5)
                ' Установить DataTable в качестве источника данных для DataGridView
                DataGridView4.DataSource = table5
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command2 As New SqlCommand("Select MZFW, MTOW, MLW  FROM [Test].[dbo].[A319] where flight_bort = @flight_bort and config_id = @config_id", connection)
                command2.Parameters.Add("@flight_bort", SqlDbType.VarChar).Value = TextBox3.Text
                command2.Parameters.Add("@config_id", SqlDbType.VarChar).Value = TextBox2.Text
                Dim adapter2 As New SqlDataAdapter(command2)
                Dim table2 As New DataTable
                adapter2.Fill(table2)
                TextBox40.Text = table2.Rows(0)(0).ToString()
                TextBox41.Text = table2.Rows(0)(1).ToString()
                TextBox42.Text = table2.Rows(0)(2).ToString()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                ' Создать команду для выборки данных
                Dim command6 As New SqlCommand("SELECT crew as Экипаж, kitchen as Кухня_тип, DOW, DOI, MAC FROM [Test].[dbo].[A319 Wt] where Flight_bort = @Flight_bort and config_id = @config_id", connection)
                command6.Parameters.AddWithValue("Flight_bort", TextBox3.Text)
                command6.Parameters.AddWithValue("config_id", TextBox2.Text)
                Dim adapter6 As New SqlDataAdapter(command6)
                Dim table6 As New DataTable
                adapter6.Fill(table6)
                ' Установить DataTable в качестве источника данных для DataGridView
                DataGridView5.DataSource = table6
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub
    Public Sub disp_data2()
        Dim connection As SqlConnection = New SqlConnection(connectionString)
        Dim command3 As New SqlCommand("SELECT FuelWeight AS Вес_топлива, CenterOfGravityIndex as Индекс FROM [Test].[dbo].[fuel_079]", connection)
        Dim adapter3 As New SqlDataAdapter(command3)
        Dim table3 As New DataTable()
        adapter3.Fill(table3)
        DataGridView2.DataSource = table3
    End Sub
    Public Sub disp_data3()
        Dim connection As SqlConnection = New SqlConnection(connectionString)
        Dim command4 As New SqlCommand("SELECT FuelWeight AS Вес_топлива, CenterOfGravityIndex as Индекс FROM [Test].[dbo].[fuel_080]", connection)
        Dim adapter4 As New SqlDataAdapter(command4)
        Dim table4 As New DataTable()
        adapter4.Fill(table4)
        DataGridView3.DataSource = table4
    End Sub
    Public Sub disp_data4()
        Dim connection As SqlConnection = New SqlConnection(connectionString)
        Dim command5 As New SqlCommand("SELECT FuelWeight AS Вес_топлива, CenterOfGravityIndex as Индекс FROM [Test].[dbo].[fuel_810]", connection)
        Dim adapter5 As New SqlDataAdapter(command5)
        Dim table5 As New DataTable()
        adapter5.Fill(table5)
        DataGridView4.DataSource = table5
    End Sub
    Public Sub disp_data5()
        Dim connection As SqlConnection = New SqlConnection(connectionString)
        Dim command6 As New SqlCommand("SELECT crew as Экипаж, kitchen as Кухня_тип, DOW, DOI, MAC FROM [Test].[dbo].[A319 Wt] where Flight_bort = @Flight_bort and config_id = @config_id", connection)
        command6.Parameters.AddWithValue("Flight_bort", TextBox3.Text)
        command6.Parameters.AddWithValue("config_id", TextBox2.Text)
        Dim adapter6 As New SqlDataAdapter(command6)
        Dim table6 As New DataTable()
        adapter6.Fill(table6)
        DataGridView5.DataSource = table6
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("UPDATE [Test].[dbo].[A319] SET [MZFW] = @MZFW, [MTOW] = @MTOW, [MLW] = @MLW WHERE flight_bort = @flight_bort and config_id = @config_id ", connection)
                command.Parameters.AddWithValue("@MZFW", TextBox40.Text)
                command.Parameters.AddWithValue("@MTOW", TextBox41.Text)
                command.Parameters.AddWithValue("@MLW", TextBox42.Text)
                command.Parameters.Add("@flight_bort", SqlDbType.VarChar).Value = TextBox3.Text
                command.Parameters.Add("@config_id", SqlDbType.VarChar).Value = TextBox2.Text
                command.ExecuteNonQuery()
                MessageBox.Show("Обновление данных")
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim selectedRow As DataGridViewRow = DataGridView2.Rows(e.RowIndex)
        Try
            TextBox45.Text = selectedRow.Cells(0).Value.ToString() ' Первый столбец
            TextBox43.Text = selectedRow.Cells(1).Value.ToString() ' Второй столбец
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("Insert into [Test].[dbo].[fuel_079] ([FuelWeight],[CenterOfGravityIndex]) Values(@FuelWeight, @CenterOfGravityIndex)", connection)
                command.Parameters.AddWithValue("@FuelWeight", TextBox45.Text)
                command.Parameters.AddWithValue("@CenterOfGravityIndex", TextBox43.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Запись в Базу Данных")
                disp_data2()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("UPDATE [Test].[dbo].[fuel_079] SET [FuelWeight] = @FuelWeight, [CenterOfGravityIndex] = @CenterOfGravityIndex WHERE FuelWeight = @FuelWeight", connection)
                command.Parameters.AddWithValue("@FuelWeight", TextBox45.Text)
                command.Parameters.AddWithValue("@CenterOfGravityIndex", TextBox43.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Обновление данных")
                disp_data2()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("DELETE FROM [Test].[dbo].[fuel_079] WHERE FuelWeight = @FuelWeight", connection)
                command.Parameters.AddWithValue("@FuelWeight", TextBox45.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Удаление данных")
                disp_data2()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Dim selectedRow As DataGridViewRow = DataGridView3.Rows(e.RowIndex)
        Try
            TextBox46.Text = selectedRow.Cells(0).Value.ToString() ' Первый столбец
            TextBox44.Text = selectedRow.Cells(1).Value.ToString() ' Второй столбец
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("Insert into [Test].[dbo].[fuel_080] ([FuelWeight],[CenterOfGravityIndex]) Values(@FuelWeight, @CenterOfGravityIndex)", connection)
                command.Parameters.AddWithValue("@FuelWeight", TextBox46.Text)
                command.Parameters.AddWithValue("@CenterOfGravityIndex", TextBox44.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Запись в Базу Данных")
                disp_data3()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("UPDATE [Test].[dbo].[fuel_080] SET [FuelWeight] = @FuelWeight, [CenterOfGravityIndex] = @CenterOfGravityIndex WHERE FuelWeight = @FuelWeight", connection)
                command.Parameters.AddWithValue("@FuelWeight", TextBox46.Text)
                command.Parameters.AddWithValue("@CenterOfGravityIndex", TextBox44.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Обновление данных")
                disp_data3()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("DELETE FROM [Test].[dbo].[fuel_080] WHERE FuelWeight = @FuelWeight", connection)
                command.Parameters.AddWithValue("@FuelWeight", TextBox46.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Удаление данных")
                disp_data3()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        Dim selectedRow As DataGridViewRow = DataGridView4.Rows(e.RowIndex)
        Try
            TextBox48.Text = selectedRow.Cells(0).Value.ToString() ' Первый столбец
            TextBox47.Text = selectedRow.Cells(1).Value.ToString() ' Второй столбец
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("Insert into [Test].[dbo].[fuel_810] ([FuelWeight],[CenterOfGravityIndex]) Values(@FuelWeight, @CenterOfGravityIndex)", connection)
                command.Parameters.AddWithValue("@FuelWeight", TextBox48.Text)
                command.Parameters.AddWithValue("@CenterOfGravityIndex", TextBox47.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Запись в Базу Данных")
                disp_data4()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("UPDATE [Test].[dbo].[fuel_810] SET [FuelWeight] = @FuelWeight, [CenterOfGravityIndex] = @CenterOfGravityIndex WHERE FuelWeight = @FuelWeight", connection)
                command.Parameters.AddWithValue("@FuelWeight", TextBox48.Text)
                command.Parameters.AddWithValue("@CenterOfGravityIndex", TextBox47.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Обновление данных")
                disp_data4()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("DELETE FROM [Test].[dbo].[fuel_810] WHERE FuelWeight = @FuelWeight", connection)
                command.Parameters.AddWithValue("@FuelWeight", TextBox48.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Удаление данных")
                disp_data4()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub DataGridView5_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellClick
        Dim selectedRow As DataGridViewRow = DataGridView5.Rows(e.RowIndex)
        Try
            TextBox50.Text = selectedRow.Cells(0).Value.ToString()
            TextBox49.Text = selectedRow.Cells(1).Value.ToString()
            TextBox51.Text = selectedRow.Cells(2).Value.ToString()
            TextBox52.Text = selectedRow.Cells(3).Value.ToString()
            TextBox53.Text = selectedRow.Cells(4).Value.ToString()
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("Insert into [Test].[dbo].[A319 Wt] ([Flight_bort],[crew],[kitchen],[DOW],[DOI],[MAC],[config_id]) Values(@Flight_bort, @crew, @kitchen, @DOW , @DOI, @MAC, @config_id)", connection)
                command.Parameters.AddWithValue("@Flight_bort", TextBox3.Text)
                command.Parameters.AddWithValue("@crew", TextBox50.Text)
                command.Parameters.AddWithValue("@kitchen", TextBox49.Text)
                command.Parameters.AddWithValue("@DOW", TextBox51.Text)
                command.Parameters.AddWithValue("@DOI", TextBox52.Text)
                command.Parameters.AddWithValue("@MAC", TextBox53.Text)
                command.Parameters.AddWithValue("@config_id", TextBox2.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Запись в Базу Данных")
                disp_data5()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("UPDATE [Test].[dbo].[A319 Wt] SET [crew] = @crew, [kitchen] = @kitchen, [DOW] = @DOW, [DOI] = @DOI, [MAC] = @MAC WHERE crew = @crew and config_id = @config_id", connection)
                command.Parameters.AddWithValue("@crew", TextBox50.Text)
                command.Parameters.AddWithValue("@kitchen", TextBox49.Text)
                command.Parameters.AddWithValue("@DOW", TextBox51.Text)
                command.Parameters.AddWithValue("@DOI", TextBox52.Text)
                command.Parameters.AddWithValue("@MAC", TextBox53.Text)
                command.Parameters.AddWithValue("@config_id", TextBox2.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Обновление данных")
                disp_data5()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("DELETE FROM [Test].[dbo].[A319 Wt] WHERE crew = @crew and config_id = @config_id", connection)
                command.Parameters.AddWithValue("@crew", TextBox50.Text)
                command.Parameters.AddWithValue("@config_id", TextBox2.Text)
                command.ExecuteNonQuery()
                MessageBox.Show("Удаление данных")
                disp_data5()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim registryKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("FormDatachanges", True)
        If registryKey Is Nothing Then
            registryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("FormDatachanges")
        End If
        ' key для form3 (MAX pax)
        registryKey.SetValue("TextBox4", TextBox4.Text)
        registryKey.SetValue("TextBox5", TextBox5.Text)
        registryKey.SetValue("TextBox6", TextBox6.Text)
        registryKey.SetValue("TextBox7", TextBox7.Text)
        'key для form3 (variants passengers)
        registryKey.SetValue("TextBox56", TextBox56.Text)
        registryKey.SetValue("TextBox39", TextBox39.Text)
        registryKey.SetValue("TextBox55", TextBox55.Text)
        registryKey.SetValue("TextBox38", TextBox39.Text)
        registryKey.SetValue("TextBox54", TextBox54.Text)
        registryKey.SetValue("TextBox37", TextBox37.Text)
        'key для form1 (MAX wt)
        registryKey.SetValue("TextBox8", TextBox8.Text)
        registryKey.SetValue("TextBox9", TextBox9.Text)
        registryKey.SetValue("TextBox10", TextBox10.Text)
        registryKey.SetValue("TextBox11", TextBox11.Text)
        'key для form1 (влияние индекса)
        registryKey.SetValue("TextBox12", TextBox12.Text)
        registryKey.SetValue("TextBox13", TextBox13.Text)
        registryKey.SetValue("TextBox14", TextBox14.Text)
        registryKey.SetValue("TextBox15", TextBox15.Text)
        'key для form1 (для вычислений)
        registryKey.SetValue("TextBox29", TextBox29.Text)
        registryKey.SetValue("TextBox30", TextBox30.Text)
        registryKey.SetValue("TextBox31", TextBox31.Text)
        registryKey.SetValue("TextBox32", TextBox32.Text)
        registryKey.SetValue("TextBox33", TextBox33.Text)
        'key для form1 (для index load)
        registryKey.SetValue("TextBox36", TextBox36.Text)
        registryKey.SetValue("TextBox35", TextBox35.Text)
        registryKey.SetValue("TextBox34", TextBox34.Text)
        'key для form1 (для max load)
        registryKey.SetValue("TextBox16", TextBox16.Text)
        registryKey.SetValue("TextBox17", TextBox17.Text)
        registryKey.SetValue("TextBox18", TextBox18.Text)
        'key для form2 (Max position)
        registryKey.SetValue("TextBox19", TextBox19.Text)
        registryKey.SetValue("TextBox20", TextBox20.Text)
        registryKey.SetValue("TextBox21", TextBox21.Text)
        registryKey.SetValue("TextBox22", TextBox22.Text)
        registryKey.SetValue("TextBox23", TextBox23.Text)
        'key для form2 (Max V position)
        registryKey.SetValue("TextBox24", TextBox24.Text)
        registryKey.SetValue("TextBox25", TextBox25.Text)
        registryKey.SetValue("TextBox26", TextBox26.Text)
        registryKey.SetValue("TextBox27", TextBox27.Text)
        registryKey.SetValue("TextBox28", TextBox28.Text)
        'изменение названий груз.отсеки
        registryKey.SetValue("Label27", Label27.Text)
        registryKey.SetValue("Label26", Label26.Text)
        registryKey.SetValue("Label25", Label25.Text)
        registryKey.Close()
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command2 As New SqlCommand("Select MZFW, MTOW, MLW  FROM [Test].[dbo].[A319] where flight_bort = @flight_bort and config_id = @config_id", connection)
                command2.Parameters.Add("@flight_bort", SqlDbType.VarChar).Value = TextBox3.Text
                command2.Parameters.Add("@config_id", SqlDbType.VarChar).Value = TextBox2.Text
                Dim adapter2 As New SqlDataAdapter(command2)
                Dim table2 As New DataTable
                adapter2.Fill(table2)
                Form1.TextBox13.Text = table2.Rows(0)(0).ToString()
                Form1.TextBox19.Text = table2.Rows(0)(1).ToString()
                Form1.TextBox18.Text = table2.Rows(0)(2).ToString()
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim command As New SqlCommand("SELECT TOP 1 * FROM [Test].[dbo].[A319 Act] where flight1_bort = @flight1_bort and type_Aircraft = @type_Aircraft ORDER BY CreatedAt DESC", connection)
                command.Parameters.Add("@flight1_bort", SqlDbType.VarChar).Value = TextBox3.Text
                command.Parameters.Add("@type_Aircraft", SqlDbType.VarChar).Value = TextBox1.Text
                Dim adapter As New SqlDataAdapter(command)
                Dim table As New DataTable
                adapter.Fill(table)
                Form1.TextBox5.Text = table.Rows(0)(0).ToString()
                Form1.TextBox3.Text = table.Rows(0)(1).ToString()
                Form1.TextBox4.Text = table.Rows(0)(2).ToString()
                Form1.TextBox2.Text = table.Rows(0)(3).ToString()
                Form1.TextBox48.Text = table.Rows(0)(4).ToString()
                Form1.TextBox49.Text = table.Rows(0)(5).ToString()
                Form1.TextBox8.Text = table.Rows(0)(6).ToString()
                Form1.TextBox12.Text = table.Rows(0)(7).ToString()
                Form1.TextBox17.Text = table.Rows(0)(8).ToString()
                Form1.TextBox20.Text = table.Rows(0)(9).ToString()
                Form1.TextBox21.Text = table.Rows(0)(10).ToString()
                Form1.TextBox42.Text = table.Rows(0)(11).ToString()
                Form1.TextBox44.Text = table.Rows(0)(12).ToString()
                Form1.TextBox46.Text = table.Rows(0)(13).ToString()
                Form1.TextBox43.Text = table.Rows(0)(14).ToString()
                Form1.TextBox45.Text = table.Rows(0)(15).ToString()
                Form1.TextBox47.Text = table.Rows(0)(16).ToString()
                Form1.TextBox6.Text = table.Rows(0)(17).ToString()
                Form1.TextBox7.Text = table.Rows(0)(18).ToString()
                Form1.TextBox14.Text = table.Rows(0)(19).ToString()
                Form1.TextBox11.Text = table.Rows(0)(19).ToString()
                Form1.ComboBox1.Text = table.Rows(0)(20).ToString()
                Form1.TextBox23.Text = table.Rows(0)(21).ToString()
                Form1.TextBox24.Text = table.Rows(0)(22).ToString()
                Form1.TextBox25.Text = table.Rows(0)(23).ToString()
                Form1.TextBox15.Text = table.Rows(0)(24).ToString()
                Form1.TextBox26.Text = table.Rows(0)(25).ToString()
                Form1.TextBox27.Text = table.Rows(0)(26).ToString()
                Form1.TextBox28.Text = table.Rows(0)(27).ToString()
                Form1.TextBox29.Text = table.Rows(0)(28).ToString()
                Form1.TextBox30.Text = table.Rows(0)(29).ToString()
                Form1.TextBox31.Text = table.Rows(0)(30).ToString()
                Form1.TextBox36.Text = table.Rows(0)(31).ToString()
                Form1.TextBox37.Text = table.Rows(0)(32).ToString()
                Form1.TextBox38.Text = table.Rows(0)(33).ToString()
                Form1.TextBox16.Text = table.Rows(0)(34).ToString()
                Form1.TextBox39.Text = table.Rows(0)(35).ToString()
                Form1.TextBox40.Text = table.Rows(0)(36).ToString()
                Form1.TextBox41.Text = table.Rows(0)(37).ToString()
                Form1.TextBox10.Text = table.Rows(0)(38).ToString()
                Form1.TextBox32.Text = table.Rows(0)(39).ToString()
                Form1.TextBox33.Text = table.Rows(0)(40).ToString()
                Form1.TextBox34.Text = table.Rows(0)(41).ToString()
                Form1.TextBox35.Text = table.Rows(0)(42).ToString()
                Form1.TextBox9.Text = table.Rows(0)(43).ToString()
                Form1.TextBox13.Text = table.Rows(0)(44).ToString()
                Form1.TextBox19.Text = table.Rows(0)(45).ToString()
                Form1.TextBox18.Text = table.Rows(0)(46).ToString()
                Form1.Label29.Text = table.Rows(0)(47).ToString()
                Form1.Label57.Text = table.Rows(0)(48).ToString()
                Form1.Label56.Text = table.Rows(0)(49).ToString()
                Form1.Label71.Text = table.Rows(0)(50).ToString()
                Form1.Label73.Text = table.Rows(0)(51).ToString()
                Form1.Label75.Text = table.Rows(0)(52).ToString()
                Form1.Label50.Text = table.Rows(0)(53).ToString()
                Form1.Label51.Text = table.Rows(0)(54).ToString()
                Form1.Label52.Text = table.Rows(0)(55).ToString()
                Form1.Label53.Text = table.Rows(0)(56).ToString()
                Form1.TextBox1.Text = table.Rows(0)(57).ToString()
            End Using
        Catch ex As Exception
            MsgBox("По данному рейсу записей не найдено!", vbInformation, "Информация")
        End Try
        Form1.TextBox5.Text = Form12.TextBox1.Text
        Form1.TextBox3.Text = Form12.TextBox2.Text
        Form1.TextBox48.Text = Form12.TextBox3.Text
        Form1.TextBox49.Text = Form12.TextBox4.Text
        Form1.TextBox1.Text = Form12.TextBox5.Text
        Form1.TextBox4.Text = Form12.TextBox6.Text
        Form1.TextBox2.Text = Form12.TextBox7.Text
        Form1.Show()
        Me.Close()
    End Sub
End Class
