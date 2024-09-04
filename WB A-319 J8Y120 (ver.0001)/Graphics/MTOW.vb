Imports SkiaSharp



Public Class MTOW
    Inherits Drawmainline
    Implements Icheckpoints

    Public Property PointsY0 As List(Of Single)
    Public Property PointsX0 As List(Of Single)
    Public Property PointsY1 As List(Of Single)
    Public Property PointsX1 As List(Of Single)
    Public Property Movex0 As List(Of Single)
    Public Property Movey0 As List(Of Single)
    Public Property Movex1 As List(Of Single)
    Public Property Movey1 As List(Of Single)
    Public Property Movex2 As List(Of Single)
    Public Property Movey2 As List(Of Single)
    Public Property Paint As SKPaint
    Public Property AllPoints As List(Of SKPoint)

    Public Sub New(canvas As SKCanvas, width As Integer, height As Integer, minX As Single, maxX As Single, minY As Single, maxY As Single, color As SKColor)
        MyBase.New(canvas, width, height, minX, maxX, minY, maxY)


        PointsY0 = New List(Of Single)()
        PointsX0 = New List(Of Single)()
        PointsY1 = New List(Of Single)()
        PointsX1 = New List(Of Single)()
        Movex0 = New List(Of Single)()
        Movey0 = New List(Of Single)()
        Movex1 = New List(Of Single)()
        Movey1 = New List(Of Single)()
        Movex2 = New List(Of Single)()
        Movey2 = New List(Of Single)()

        AllPoints = New List(Of SKPoint)()

        Me.Paint = New SKPaint With {
           .Color = color,
           .IsAntialias = True,
           .Style = SKPaintStyle.Stroke,
           .StrokeWidth = 2
       }

    End Sub

    ' Метод для добавления линии и сохранения точек
    Public Sub AddLine(x0 As Single, y0 As Single, x1 As Single, y1 As Single)
        PointsX0.Add(x0)
        PointsY0.Add(y0)
        PointsX1.Add(x1)
        PointsY1.Add(y1)

        ' Сохраняем точки в общий список
        If Not AllPoints.Contains(New SKPoint(x0, y0)) Then
            AllPoints.Add(New SKPoint(x0, y0))
        End If
        If Not AllPoints.Contains(New SKPoint(x1, y1)) Then
            AllPoints.Add(New SKPoint(x1, y1))
        End If

    End Sub

    ' Метод для добавления кривой и сохранения точек
    Public Sub AddCurve(movex0 As Single, movey0 As Single, movex1 As Single, movey1 As Single, movex2 As Single, movey2 As Single)
        Me.Movex0.Add(movex0)
        Me.Movey0.Add(movey0)
        Me.Movex1.Add(movex1)
        Me.Movey1.Add(movey1)
        Me.Movex2.Add(movex2)
        Me.Movey2.Add(movey2)

        ' Сохраняем контрольные точки кривой в общий список
        If Not AllPoints.Contains(New SKPoint(movex0, movey0)) Then
            AllPoints.Add(New SKPoint(movex0, movey0))
        End If
        If Not AllPoints.Contains(New SKPoint(movex1, movey1)) Then
            AllPoints.Add(New SKPoint(movex1, movey1))
        End If
        If Not AllPoints.Contains(New SKPoint(movex2, movey2)) Then
            AllPoints.Add(New SKPoint(movex2, movey2))
        End If
    End Sub

    ' Метод для рисования всех линий
    Public Sub DrawAllLines()
        For i = 0 To Math.Min(PointsX0.Count, PointsY0.Count) - 1
            DrawLine(Paint, PointsX0(i), PointsY0(i), PointsX1(i), PointsY1(i))
        Next
    End Sub

    ' Метод для рисования всех кривых
    Public Sub DrawAllCurves()

        For i = 0 To Math.Min(Movex0.Count, Movey0.Count) - 1
            Drawcurves(Paint, Movex0(i), Movey0(i), Movex1(i), Movey1(i), Movex2(i), Movey2(i))
        Next
    End Sub

    ' Метод для проверки, находится ли точка внутри многоугольника
    Public Function ContainsPoint(x As Single, y As Single) As Boolean Implements Icheckpoints.ContainsPoint
        Dim point As New SKPoint(x, y)
        Dim result As Boolean = False
        Dim j As Integer = AllPoints.Count - 1
        For i As Integer = 0 To AllPoints.Count - 1
            If (AllPoints(i).Y < point.Y AndAlso AllPoints(j).Y >= point.Y OrElse AllPoints(j).Y < point.Y AndAlso AllPoints(i).Y >= point.Y) AndAlso
               (AllPoints(i).X + (point.Y - AllPoints(i).Y) / (AllPoints(j).Y - AllPoints(i).Y) * (AllPoints(j).X - AllPoints(i).X) < point.X) Then
                result = Not result
            End If
            j = i
        Next
        Return result
    End Function

    Public Sub UpdateColor(newColor As SKColor)
        Me.Paint.Color = newColor
    End Sub

    Public Sub ClearList()
        PointsY0.Clear()
        PointsX0.Clear()
        PointsY1.Clear()
        PointsX1.Clear()
        Movex0.Clear()
        Movey0.Clear()
        Movex1.Clear()
        Movey1.Clear()
        Movex2.Clear()
        Movey2.Clear()
    End Sub

End Class
