Public Class GroupData

    Public Property LableCAX As String
    Public Property CAX As Single
    Public Property Index As Single
    Public Property Weight As Integer

    Public Sub New(LableCAX As String, CAX As Single, Index As Single, Weight As Integer)
        Me.LableCAX = LableCAX
        Me.CAX = CAX
        Me.Index = Index
        Me.Weight = Weight
    End Sub
End Class
