Imports System.Xml


Public Class DatabaseConnections
    ' это модификатор доступа/только внутри класса
    Private connectionStrings As Dictionary(Of String, String)

    Public Sub New()
        'тут происходит инициализация словаря и инициализация xml
        connectionStrings = New Dictionary(Of String, String)()
        LoadConnectionStringsFromXml("dbconnect.xml")
    End Sub
    Private Sub LoadConnectionStringsFromXml(ByVal xmlFilePath As String)
        Try
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(xmlFilePath)

            Dim accessNode As XmlNode = xmlDoc.SelectSingleNode("/database/stringconect_access")
            Dim mainNode As XmlNode = xmlDoc.SelectSingleNode("/database/stringconect_main")

            If accessNode IsNot Nothing Then
                connectionStrings.Add("stringconect_access", accessNode.InnerText)
            End If

            If mainNode IsNot Nothing Then
                connectionStrings.Add("stringconect_main", mainNode.InnerText)
            End If
        Catch ex As Exception
            Console.WriteLine($"Error load xml:{ex}")
        End Try
    End Sub

    Public Function GetConnectionString(ByVal key As String) As String
        If connectionStrings.ContainsKey(key) Then
            Return connectionStrings(key)
        Else
            Return Nothing ' или можно выбрасывать исключение, если ключ не найден
        End If
    End Function

End Class