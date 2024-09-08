Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Reflection
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
            Dim connectionStr As String = GetConnectionString()
            Using connection As New SqlConnection(connectionStr)

                Dim username As String = UsernameTextBox.Text
                Dim password As String = PasswordTextBox.Text
                If CheckCredentials(connection, username, password) Then
                    ShowAccessGranted()
                Else
                    ShowUserNotFound()
                End If
            End Using
        Catch ex As Exception
            HandleError(ex)
        End Try
    End Sub

    Private Function GetConnectionString() As String
        Dim dbconnections As New DatabaseConnections()
        Dim sdaccessConnectionString As String = dbconnections.GetConnectionString("stringconect_access")
        log.Info($"Получили строку подключения {sdaccessConnectionString}")
        Return sdaccessConnectionString
    End Function

    Private Function CheckCredentials(connection As SqlConnection, username As String, password As String) As Boolean
        Using command As New SqlCommand("CheckCredentials", connection)
            command.CommandType = CommandType.StoredProcedure
            command.Parameters.Add("@username", SqlDbType.VarChar).Value = username
            command.Parameters.Add("@password", SqlDbType.VarChar).Value = password
            Using dt As New DataTable
                Using adapter As New SqlDataAdapter(command)
                    adapter.Fill(dt)
                    Return dt.Rows.Count > 0
                End Using
            End Using
        End Using
    End Function

    Private Sub ShowAccessGranted()
        MsgBox("Доступ разрешён!", MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
        log.Info("Доступ разрешён!")
        Form12.Show()
        Me.Close()
    End Sub

    Private Sub ShowUserNotFound()
        MsgBox("Пользователь не найден!", MsgBoxStyle.OkOnly + MsgBoxStyle.Critical)
        log.Warn("Пользователь не найден!")
        UsernameTextBox.Text = ""
        PasswordTextBox.Text = ""
    End Sub

    Private Sub HandleError(ex As Exception)
        MsgBox("Error: " & ex.ToString())
        log.Error("Error" & ex.ToString)
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        log.Warn("Closed")
        Me.Close()
    End Sub


End Class
