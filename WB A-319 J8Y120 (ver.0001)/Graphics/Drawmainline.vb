Imports SkiaSharp

Public Class Drawmainline
    Inherits DrawLine

    Public Sub New(canvas As SKCanvas, width As Integer, height As Integer, minX As Integer, maxX As Integer, minY As Integer, maxY As Integer)

        MyBase.New(canvas, width, height, minX, maxX, minY, maxY)

    End Sub

    Public Sub Drawcurves(paint As SKPaint, movex0 As Single, movey0 As Single, movex1 As Single, movey1 As Single, movex2 As Single, movey2 As Single)

        Using path As New SKPath()
            path.MoveTo(Xval(movex0), Yval(movey0))
            path.QuadTo(Xval(movex1), Yval(movey1), Xval(movex2), Yval(movey2))
            Canvas.DrawPath(path, paint)
        End Using
    End Sub

    Public Sub DrawLine(linepaint As SKPaint, x0 As Single, y0 As Single, x1 As Single, y1 As Single)


        Canvas.DrawLine(Xval(x0), Yval(y0), Xval(x1), Yval(y1), linepaint)


    End Sub

End Class
