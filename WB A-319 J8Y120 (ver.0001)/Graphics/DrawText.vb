Imports SkiaSharp

Public Class DrawText
    Inherits DrawLine
    Implements IUpcast

    Public Sub New(canvas As SKCanvas, width As Integer, height As Integer, minX As Integer, maxX As Integer, minY As Integer, maxY As Integer)
        MyBase.New(canvas, width, height, minX, maxX, minY, maxY) ' Вызов конструктора базового класса
    End Sub


    Public Overloads Sub DrawVerticalText() Implements IUpcast.DrawVerticalText
        Canvas.Save()
        Canvas.RotateDegrees(-90, StartX, StartY)
        'Canvas.DrawText("Aircraft Weight (kg x 1000)", StartX + 230, StartY - 50, AxisPaint)
        Canvas.DrawText("Aircraft Weight (kg x 1000)", StartX + 270, StartY - 50, AxisPaint)
        Canvas.Restore()
        Canvas.DrawText("Index", StartX - 45, StartY + 20, AxisPaint)
        'Canvas.DrawText("Aircraft CG (%MAC)", StartX + 950, StartY - 705, AxisPaint)
        Canvas.DrawText("Aircraft CG (%MAC)", StartX + 950, StartY - 685, AxisPaint)
    End Sub


End Class
