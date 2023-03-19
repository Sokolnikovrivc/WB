Imports System.Data.SqlClient
Imports System.Web

Public Class Form5
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim ost As Single = 0
        For Each val As String In TextBox1.Text.Split(" "c)
            Dim num As Single
            If Single.TryParse(val, num) Then
                ost += num
            End If
        Next
        TextBox5.Text = (ost + GetZapValue() + GetRulValue()).ToString()
        TextBox7.Text = (ost + GetZapValue() + GetRulValue()).ToString()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim zap As Single = 0
        For Each val As String In TextBox2.Text.Split(" "c)
            Dim num As Single
            If Single.TryParse(val, num) Then
                zap += num
            End If
        Next
        TextBox5.Text = (zap + GetOstValue() + GetRulValue()).ToString()
        TextBox7.Text = (zap + GetOstValue() + GetRulValue()).ToString()
    End Sub

    Private Function GetZapValue() As Single
        Dim zap As Single = 0
        For Each val As String In TextBox2.Text.Split(" "c)
            Dim num As Single
            If Single.TryParse(val, num) Then
                zap += num
            End If
        Next
        Return zap
    End Function

    Private Function GetOstValue() As Single
        Dim ost As Single = 0
        For Each val As String In TextBox1.Text.Split(" "c)
            Dim num As Single
            If Single.TryParse(val, num) Then
                ost += num
            End If
        Next
        Return ost
    End Function

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Dim rul As Single = 0
        For Each val As String In TextBox6.Text.Split(" "c)
            Dim num As Single
            If Single.TryParse(val, num) Then
                rul += num
            End If
        Next
        TextBox5.Text = (rul + GetZapValue() + GetOstValue()).ToString()
        TextBox7.Text = (rul + GetZapValue() + GetOstValue()).ToString()
    End Sub
    Private Function GetRulValue() As Single
        Dim rul As Single = 0
        For Each val As String In TextBox6.Text.Split(" "c)
            Dim num As Single
            If Single.TryParse(val, num) Then
                rul += num
            End If
        Next
        Return rul
    End Function
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim connectionString As String = "Data Source=WIN-8CEIKSU78CS\SQLEXPRESS; Initial Catalog=Test; Integrated Security=True"
        Dim fuelWeight As Integer = CInt(TextBox5.Text)
        Dim fuelWeight2 As Integer = CInt(TextBox4.Text)

        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()

                Dim query As String = ""
                Select Case ComboBox1.SelectedIndex
                    Case 0
                        query = "SELECT TOP 1 CenterOfGravityIndex FROM fuel_079 ORDER BY ABS(FuelWeight - @fuelWeight) ASC"
                    Case 1
                        query = "SELECT TOP 1 CenterOfGravityIndex FROM fuel_080 ORDER BY ABS(FuelWeight - @fuelWeight) ASC"
                    Case 2
                        query = "SELECT TOP 1 CenterOfGravityIndex FROM fuel_810 ORDER BY ABS(FuelWeight - @fuelWeight) ASC"
                End Select

                Dim command As SqlCommand = New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@fuelWeight", fuelWeight)

                Dim result As Object = command.ExecuteScalar()
                If result IsNot Nothing Then
                    Form1.TextBox26.Text = result
                Else
                    MsgBox("Error!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
                End If
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
        Try
            Using connection As SqlConnection = New SqlConnection(connectionString)
                connection.Open()
                Dim query As String = ""
                Select Case ComboBox1.SelectedIndex
                    Case 0
                        query = "SELECT TOP 1 CenterOfGravityIndex FROM fuel_079 ORDER BY ABS(FuelWeight - @fuelWeight2) ASC"
                    Case 1
                        query = "SELECT TOP 1 CenterOfGravityIndex FROM fuel_080 ORDER BY ABS(FuelWeight - @fuelWeight2) ASC"
                    Case 2
                        query = "SELECT TOP 1 CenterOfGravityIndex FROM fuel_810 ORDER BY ABS(FuelWeight - @fuelWeight2) ASC"
                End Select

                ' Поиск ближайшего значения индекса для заданного веса топлива
                Dim command As SqlCommand = New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@fuelWeight2", fuelWeight2)

                Dim result As Object = command.ExecuteScalar()

                If result IsNot Nothing Then
                    Form1.TextBox27.Text = result
                Else
                    MsgBox("Error!", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical, "Ошибка!")
                End If
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.TextBox15.Text = TextBox5.Text
        Form1.TextBox20.Text = TextBox4.Text
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = "0"
        TextBox2.Text = "0"
        TextBox4.Text = "0"
        TextBox5.Text = "0"
        TextBox6.Text = "0"
        TextBox7.Text = "0"
        Me.Close()
    End Sub
End Class