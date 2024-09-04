Imports SkiaSharp

Public Class CAX
    Inherits DrawLine

    Public Property Data As Dictionary(Of Single, Single)
    Public Property ClosestValues As Dictionary(Of Single, Single)
    Public Property CAXHead As Dictionary(Of Single, Integer)
    Public Property Indexbottom As List(Of Single)


    Public Sub New(canvas As SKCanvas, width As Integer, height As Integer, minX As Single, maxX As Single, minY As Single, maxY As Single)
        MyBase.New(canvas, width, height, minX, maxX, minY, maxY)
        Me.Data = New Dictionary(Of Single, Single)()
        Me.ClosestValues = New Dictionary(Of Single, Single)()
        Me.CAXHead = New Dictionary(Of Single, Integer)()
        Me.Indexbottom = New List(Of Single)()

    End Sub

    Public Sub DrawLineCAX(С As Integer, CAX As Single, LEMAC As Single, RefSTA As Integer, K As Integer)


        For i As Single = MinX To MaxX Step 0.1
                Dim CAXorig As Single
                Dim W As Integer = MaxY * 1000


                CAXorig = ((((С * (i - K)) / W) + RefSTA - LEMAC) / CAX) * 100
                Data.Add(i, CAXorig)

            Next

            For Each kvp In Data

                Dim value As Single = kvp.Value
                Dim roundedValue As Single = Math.Round(value)

                If Not ClosestValues.ContainsKey(roundedValue) OrElse DistanceToNearestInteger(value) < DistanceToNearestInteger(ClosestValues(roundedValue)) Then
                    ClosestValues(roundedValue) = value
                End If
            Next
            For Each kvp In ClosestValues
                Dim index As Single = Data.First(Function(x) x.Value = kvp.Value).Key
                'MsgBox($"Итерация: {index}, Значение: {kvp.Value}, Округленное значение: {Math.Round(kvp.Value)}")
                CAXHead.Add(index, Math.Round(kvp.Value))
                'MsgBox($"Значение индекса {index} и САХ: {Math.Round(kvp.Value)}")
            Next
            For Each kvp In CAXHead
                Dim CAXorig As Integer = kvp.Value
                Dim W As Integer = MinY * 1000
                Dim i As Single
                i = (((((CAXorig * CAX) / 100) - RefSTA + LEMAC) * W) / С) + K
                Indexbottom.Add(i)
                'MsgBox($"Индексы x0:{i}")
                'MsgBox($"{kvp.Key}")
            Next

            Dim minLength As Integer = Math.Min(CAXHead.Count, Indexbottom.Count)
            Dim CAXHeadList As List(Of KeyValuePair(Of Single, Integer)) = CAXHead.ToList()
            For i As Integer = 0 To minLength - 1
                Dim drowCAX = CAXHeadList(i)
                Dim y0 = Yval(MinY)
                Dim y1 = Yval(MaxY)
                Dim x1 = Xval(drowCAX.Key)
                Dim x0 = Xval(Indexbottom(i))

                Canvas.DrawLine(x0, y0, x1, y1, CAXGrid)
                Canvas.DrawLine(Xval(MinX), Yval(MaxY), Xval(MaxX), Yval(MaxY), AxisPaint)
                Canvas.DrawText(drowCAX.Value.ToString, x1 - 5, y1 - 5, AxisPaint)
            Next

    End Sub

    Private Function DistanceToNearestInteger(value As Single) As Single
        Return Math.Abs(value - Math.Round(value))
    End Function
End Class
