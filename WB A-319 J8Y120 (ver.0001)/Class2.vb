Imports System.Xml
Public Class DatabaseConfigLoader
        Private ReadOnly xmlFilePath As String

        Public Sub New(filePath As String)
            xmlFilePath = filePath
        End Sub
    ' получаем в этой конструкции путь к xml и сохраняем
    Public Function GetConnectionString() As String
            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(xmlFilePath)

        Dim dbConfig As New DataBaseConfig()
        dbConfig.StringPath = xmlDoc.SelectSingleNode("/database/stringconect").InnerText
        Dim connectionString As String = String.Format("Data Source={0}", dbConfig.StringPath)
        Return connectionString
    End Function

End Class

