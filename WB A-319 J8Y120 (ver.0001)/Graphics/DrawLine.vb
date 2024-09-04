Imports SkiaSharp

Public Class DrawLine
    Implements IDrawable
    Implements IDrawHandler

    Public Property Canvas As SKCanvas
    Public Property Width As Integer
    Public Property Height As Integer
    Public Property StartX As Integer
    Public Property StartY As Integer
    Public Property ScaleX As Single
    Public Property ScaleY As Single
    Public Property ScaleYgrid As Single
    Public Property ScaleXgrid As Single
    Public Property AxisPaint As SKPaint
    Public Property GridPaint As SKPaint
    Public Property CAXGrid As SKPaint
    Public Property TextPaint As SKPaint
    Public Property MinX As Single
    Public Property MinY As Single
    Public Property MaxX As Single
    Public Property MaxY As Single
    Public Property Upcast As IUpcast


    Public Sub New(canvas As SKCanvas, width As Integer, height As Integer, minX As Single, maxX As Single, minY As Single, maxY As Single)
        Me.Canvas = canvas
        Me.Width = width
        Me.Height = height
        Me.StartY = height - 60
        Me.StartX = 100
        Me.MinX = minX
        Me.MinY = minY
        Me.MaxY = maxY
        Me.MaxX = maxX

        Dim scales = Mashtab(minX, maxX, minY, maxY)
        Dim grid = Mashtabgrid(minX, maxX, minY, maxY)
        Me.ScaleX = scales("X")
        Me.ScaleY = scales("Y")
        Me.ScaleXgrid = grid("X")
        Me.ScaleYgrid = grid("Y")

        Me.AxisPaint = New SKPaint With {
           .Color = SKColors.Black,
           .IsAntialias = True,
           .Style = SKPaintStyle.Stroke,
           .StrokeWidth = 1
       }
        Me.GridPaint = New SKPaint With {
            .Color = SKColors.Gray,
            .IsAntialias = True,
            .Style = SKPaintStyle.Stroke,
            .StrokeWidth = 1
        }
        Me.CAXGrid = New SKPaint With {
            .Color = SKColors.Gray,
            .IsAntialias = True,
            .Style = SKPaintStyle.Stroke,
            .StrokeWidth = 1.5
          }
        Me.TextPaint = New SKPaint With {
         .Color = SKColors.Black,
         .IsAntialias = True,
         .Style = SKPaintStyle.Fill,
         .StrokeWidth = 1.5
       }

    End Sub

    Public Sub DrawLineXY() Implements IDrawable.DrawLineXY
        Canvas.DrawLine(StartX, StartY, Width - 30, StartY, AxisPaint)
        Canvas.DrawLine(StartX, StartY, StartX, 30, AxisPaint)
    End Sub

    Public Sub DrawGrid() Implements IDrawable.DrawGrid

        'For i = 0 To ScaleXgrid
        '    Dim x = StartX + i * ScaleX
        '    Canvas.DrawLine(x, StartY, x, 30, GridPaint)
        'Next

        'For i = 0 To ScaleYgrid
        '    Dim y = StartY - i * ScaleY
        '    Canvas.DrawLine(StartX, y, Width - 30, y, GridPaint)
        'Next

        Canvas.DrawLine(Width - 30, StartY, Width - 30, 30, AxisPaint)
    End Sub

    'Public Sub DrawVerticalText() Implements IDrawable.DrawVerticalText

    'End Sub

    Public Sub DrawDataLineXY() Implements IDrawable.DrawDataLineXY
        'Dim valuesY As Single() = New Single() {40, 45, 50, 55, 60, 65, 70, 75, 80}
        'Dim valuesX As Single() = New Single() {25, 30, 35, 40, 45, 50, 55, 60, 65, 70, 75, 80, 85, 90, 95, 100, 105}
        Dim valuesX As New List(Of Integer)(SerchDataline(MinX, MaxX))


        ' Рисуем метки оси X
        For Each xValue In valuesX
            Dim x = Xval(xValue)
            Canvas.DrawLine(x, StartY, x, StartY + 5, AxisPaint) ' Короткая линия под меткой
            Canvas.DrawText(xValue.ToString(), x - 5, StartY + 20, AxisPaint) ' Текст метки
        Next
        valuesX.Clear()
        Dim valuesY As New List(Of Integer)(SerchDataline(MinY, MaxY))
        ' Рисуем метки оси Y
        For Each yValue In valuesY
            Dim y = Yval(yValue)
            Canvas.DrawLine(StartX, y, Xval(MaxX), y, AxisPaint)
            Canvas.DrawLine(StartX - 5, y, StartX, y, AxisPaint) ' Короткая линия рядом с меткой
            Canvas.DrawText(yValue.ToString(), StartX - 30, y + 5, AxisPaint) ' Текст метки
        Next
    End Sub

    Private Shared Function SerchDataline(min As Integer, max As Integer) As List(Of Integer)
        Dim data As New List(Of Integer)
        For i = min To max
            If i Mod 5 = 0 Then
                data.Add(i)
            End If
        Next
        Return data
    End Function

    Public Function Xval(x As Single) As Single Implements IDrawHandler.Xval
        Return StartX + (x - MinX) * ScaleX
    End Function

    Public Function Yval(y As Single) As Single Implements IDrawHandler.Yval
        Return StartY - (y - MinY) * ScaleY
    End Function

    Private Shared Function Mashtabgrid(minX As Integer, maxX As Integer, minY As Integer, maxY As Integer) As Dictionary(Of String, Integer)
        Dim datamashgrid As New Dictionary(Of String, Integer)
        Dim keyX As String = "X"
        Dim keyY As String = "Y"
        Dim resultX As Integer
        Dim resultY As Integer

        resultX = maxX - minX
        resultY = maxY - minY

        datamashgrid.Add(keyX, resultX)
        datamashgrid.Add(keyY, resultY)
        Return datamashgrid
    End Function

    Private Function Mashtab(minX As Integer, maxX As Integer, minY As Integer, maxY As Integer) As Dictionary(Of String, Double)
        Dim datamash As New Dictionary(Of String, Double)
        Dim keyX As String = "X"
        Dim keyY As String = "Y"
        Dim resultX As Double
        Dim resultY As Double

        resultX = (Width - StartX - 30) / (maxX - minX)
        resultY = (StartY - 30) / (maxY - minY)

        datamash.Add(keyX, resultX)
        datamash.Add(keyY, resultY)
        Return datamash

    End Function

    Public Sub UpdateCanvas(newCanvas As SKCanvas)
        Me.Canvas = newCanvas
    End Sub

    Public Sub UpdateSize(newWidth As Integer, newHeight As Integer)
        Me.Width = newWidth
        Me.Height = newHeight
    End Sub
End Class
