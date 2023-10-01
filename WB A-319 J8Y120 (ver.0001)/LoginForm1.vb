Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports log4net
Imports log4net.Config

' Загрузка конфигурации log4net из файла log4net.config
<Assembly: XmlConfigurator(Watch:=True)>
Public Class LoginForm1
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(LoginForm1))
    'Private Shared ReadOnly log As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Try
            Dim configLoader As New DatabaseConfigLoader("dbconnect.xml")
            Dim connectionstr As String = configLoader.GetConnectionString()
            Using connection As New SqlConnection(connectionstr)
                ' Dim connection As New SqlConnection(connectionstr)
                'New SqlConnection("Data Source=DESKTOP-7G0KEVG\SQLEXPRESS;Initial Catalog=Access;Integrated Security=True")
                Dim command As New SqlCommand("Select * from [Access].[dbo].[Control] where username = @username and password = @password", connection)
                command.Parameters.Add("@username", SqlDbType.VarChar).Value = UsernameTextBox.Text
                command.Parameters.Add("@password", SqlDbType.VarChar).Value = PasswordTextBox.Text
                Dim dt As New SqlDataAdapter(command)
                Dim table As New DataTable
                dt.Fill(table)
                If (table.Rows.Count) > 0 Then
                    MsgBox("Доступ разрешён!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
                    log.Info("Доступ разрешён!")
                    Form12.Show()
                    Me.Close()
                ElseIf (table.Rows.Count) = 0 Then
                    MsgBox("Пользователь не найден!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical)
                    log.Warn("Пользователь не найден!")
                    UsernameTextBox.Text = ""
                    PasswordTextBox.Text = ""
                End If
            End Using
        Catch ex As Exception
            MsgBox("Error: " & ex.ToString())
            log.Error("Error" & ex.ToString)
        End Try
    End Sub
    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        log.Warn("Closed")
        Me.Close()
    End Sub
End Class
