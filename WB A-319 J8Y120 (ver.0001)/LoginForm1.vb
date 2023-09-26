Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class LoginForm1
    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim configLoader As New DatabaseConfigLoader("dbconnect.xml")
        Dim connectionstr As String = configLoader.GetConnectionString()
        Dim connection As New SqlConnection(connectionstr)
        'New SqlConnection("Data Source=DESKTOP-7G0KEVG\SQLEXPRESS;Initial Catalog=Access;Integrated Security=True")
        Dim command As New SqlCommand("Select * from [Access].[dbo].[Control] where username = @username and password = @password", connection)
        command.Parameters.Add("@username", SqlDbType.VarChar).Value = UsernameTextBox.Text
        command.Parameters.Add("@password", SqlDbType.VarChar).Value = PasswordTextBox.Text
        Dim dt As New SqlDataAdapter(command)
        Dim table As New DataTable
        dt.Fill(table)
        If (table.Rows.Count) > 0 Then
            MsgBox("Доступ разрешён!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
            Form12.Show()
            Me.Close()
        ElseIf (table.Rows.Count) = 0 Then
            MsgBox("Пользователь не найден!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical)
            UsernameTextBox.Text = ""
            PasswordTextBox.Text = ""
        End If
    End Sub
    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

End Class
