Imports System.Data.SqlClient
Imports System.Reflection.Metadata
Imports OpenTK.Input
Imports SkiaSharp
Imports SkiaSharp.Views.Desktop

Public Class Form16

    Private WithEvents SkControl As SKControl
    Private MtowLine As List(Of Tuple(Of Single, Single, Single, Single))
    Private MtowCurve As List(Of Tuple(Of Single, Single, Single, Single, Single, Single))
    Private MlwLine As List(Of Tuple(Of Single, Single, Single, Single))
    Private MlwCurve As List(Of Tuple(Of Single, Single, Single, Single, Single, Single))
    Private MzfwLine As List(Of Tuple(Of Single, Single, Single, Single))
    Private MzfwCurve As List(Of Tuple(Of Single, Single, Single, Single, Single, Single))
    Private dbconnections As New DatabaseConnections()
    Private connectionstr As String = dbconnections.GetConnectionString("stringconect_main")
    Private chartDrawer As DrawLine
    Private mtow As MTOW
    Private mlw As MLW
    Private mzfw As MZFW
    Private minX As Single
    Private maxX As Single
    Private minY As Single
    Private maxY As Single
    Private c As Integer
    Private _cax As Single
    Private lemac As Single
    Private refsta As Single
    Private k As Integer
    Private colormtow As SKColor
    Private colormlw As SKColor
    Private colorzfw As SKColor

    Public Sub New()

        InitializeComponent()
        SkControl = New SKControl() With {
            .Dock = DockStyle.Fill
        }
        Controls.Add(SkControl)
        AddHandler SkControl1.PaintSurface, AddressOf OnPaintSurface
        MtowLine = New List(Of Tuple(Of Single, Single, Single, Single))()
        MtowCurve = New List(Of Tuple(Of Single, Single, Single, Single, Single, Single))()
        MlwLine = New List(Of Tuple(Of Single, Single, Single, Single))()
        MlwCurve = New List(Of Tuple(Of Single, Single, Single, Single, Single, Single))()
        MzfwLine = New List(Of Tuple(Of Single, Single, Single, Single))()
        MzfwCurve = New List(Of Tuple(Of Single, Single, Single, Single, Single, Single))()

    End Sub

    Private Sub OnPaintSurface(sender As Object, e As SKPaintSurfaceEventArgs)

        Dim canvas = e.Surface.Canvas
        Dim width = e.Info.Width
        Dim height = e.Info.Height

        canvas.Clear(SKColors.White)
        mtow.UpdateCanvas(canvas)
        mlw.UpdateCanvas(canvas)
        mzfw.UpdateCanvas(canvas)

        chartDrawer = New DrawLine(canvas, width, height, minX, maxX, minY, maxY)
        chartDrawer.DrawDataLineXY()
        chartDrawer.DrawLineXY()
        chartDrawer.DrawGrid()
        Dim CAX As New CAX(canvas, width, height, minX, maxX, minY, maxY)
        CAX.DrawLineCAX(c, _cax, lemac, refsta, k)

        chartDrawer.Upcast = New DrawText(canvas, width, height, minX, maxX, minY, maxY)
        chartDrawer.Upcast.DrawVerticalText()

        If MtowLine.Count > 0 Then
            For i = 0 To MtowLine.Count - 1
                Dim line = MtowLine(i)
                mtow.AddLine(line.Item1, line.Item2, line.Item3, line.Item4)
            Next
        End If

        If MtowCurve.Count > 0 Then
            For i = 0 To MtowCurve.Count - 1
                Dim curv = MtowCurve(i)
                mtow.AddCurve(curv.Item1, curv.Item2, curv.Item3, curv.Item4, curv.Item5, curv.Item6)
            Next
        End If
        mtow.DrawAllCurves()
        mtow.DrawAllLines()

        If MlwLine.Count > 0 Then
            For i = 0 To MlwLine.Count - 1
                Dim line = MlwLine(i)
                mlw.AddLine(line.Item1, line.Item2, line.Item3, line.Item4)
            Next
        End If

        If MlwCurve.Count > 0 Then
            For i = 0 To MlwCurve.Count - 1
                Dim curv = MlwCurve(i)
                mlw.AddCurve(curv.Item1, curv.Item2, curv.Item3, curv.Item4, curv.Item5, curv.Item6)
            Next
        End If

        mlw.DrawAllCurves()
        mlw.DrawAllLines()

        If MzfwLine.Count > 0 Then
            For i = 0 To MzfwLine.Count - 1
                Dim line = MzfwLine(i)
                mzfw.AddLine(line.Item1, line.Item2, line.Item3, line.Item4)
            Next
        End If

        If MzfwCurve.Count > 0 Then
            For i = 0 To MzfwCurve.Count - 1
                Dim curv = MzfwCurve(i)
                mzfw.AddCurve(curv.Item1, curv.Item2, curv.Item3, curv.Item4, curv.Item5, curv.Item6)
            Next
        End If

        mzfw.DrawAllCurves()
        mzfw.DrawAllLines()

        '=================================================================================================================================

        canvas.DrawCircle(chartDrawer.Xval(Single.Parse(TextBox6.Text)), chartDrawer.Yval(Single.Parse(TextBox9.Text) / 1000), 2, mtow.Paint)
        canvas.DrawCircle(chartDrawer.Xval(Single.Parse(TextBox4.Text)), chartDrawer.Yval(Single.Parse(TextBox7.Text) / 1000), 2, mlw.Paint)
        canvas.DrawCircle(chartDrawer.Xval(Single.Parse(TextBox5.Text)), chartDrawer.Yval(Single.Parse(TextBox8.Text) / 1000), 2, mzfw.Paint)


        canvas.DrawText($"TOW:({(TextBox9.Text)}; {TextBox6.Text})", chartDrawer.Xval(Single.Parse(TextBox6.Text)) - 20, chartDrawer.Yval(Single.Parse(TextBox9.Text) / 1000) - 10, chartDrawer.TextPaint)
        canvas.DrawText($"LW:({(TextBox7.Text)}; {TextBox4.Text})", chartDrawer.Xval(Single.Parse(TextBox4.Text)) - 20, chartDrawer.Yval(Single.Parse(TextBox7.Text) / 1000) - 10, chartDrawer.TextPaint)
        canvas.DrawText($"ZFW:({(TextBox8.Text)}; {TextBox5.Text})", chartDrawer.Xval(Single.Parse(TextBox5.Text)) - 20, chartDrawer.Yval(Single.Parse(TextBox8.Text) / 1000) - 10, chartDrawer.TextPaint)


        If mtow.ContainsPoint(Single.Parse(TextBox6.Text), Single.Parse(TextBox9.Text) / 1000) = False Then
            'canvas.DrawText("Точка не входит в ЦГ", chartDrawer.Xval(Single.Parse(TextBox6.Text)) + 2, chartDrawer.Yval(Single.Parse(TextBox9.Text) / 1000) + 2, chartDrawer.AxisPaint)
            MsgBox($"Точка TOW  с значениями веса: {(TextBox9.Text)} и индекса {TextBox6.Text} не входит в ЦГ", vbCritical)
            Label22.Text = "Ошибка контрольных точек ЦГ!"
            Label23.Visible = True
        End If
        If mlw.ContainsPoint(Single.Parse(TextBox4.Text), Single.Parse(TextBox7.Text) / 1000) = False Then
            'canvas.DrawText("Точка не входит в ЦГ", chartDrawer.Xval(Single.Parse(TextBox4.Text)) + 2, chartDrawer.Yval(Single.Parse(TextBox7.Text) / 1000) + 2, chartDrawer.AxisPaint)
            MsgBox($"Точка LW  с значениями веса: {(TextBox7.Text)} и индекса {TextBox4.Text} не входит в ЦГ", vbCritical)
            Label22.Text = "Ошибка контрольных точек ЦГ!"
            Label24.Visible = True
        End If
        If mzfw.ContainsPoint(Single.Parse(TextBox5.Text), Single.Parse(TextBox8.Text) / 1000) = False Then
            'canvas.DrawText("Точка не входит в ЦГ", chartDrawer.Xval(Single.Parse(TextBox5.Text)) + 2, chartDrawer.Yval(Single.Parse(TextBox8.Text) / 1000) + 2, chartDrawer.AxisPaint)
            MsgBox($"Точка ZFW  с значениями веса: {(TextBox8.Text)} и индекса {TextBox5.Text} не входит в ЦГ", vbCritical)
            Label22.Text = "Ошибка контрольных точек ЦГ!"
            Label25.Visible = True
        End If
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Using connection As New SqlConnection(connectionstr)
                connection.Open()
                Dim query As String = "
            WITH RankedData AS (
                SELECT 
                    DL.Id AS DataLineId,
                    GI.Id AS GraphicsItemId,
                    DL.NameLine,
                    DL.CreatedAt AS DL_CreatedAt,
                    GI.CreatedAt AS GI_CreatedAt,
                    DL.AircraftType,
                    DL.flight_bort,
                    DL.PointsX0,
                    DL.PointsY0,
                    DL.PointsX1,
                    DL.PointsY1,
                    DL.Movex0,
                    DL.Movey0,
                    DL.Movex1,
                    DL.Movey1,
                    DL.Movex2,
                    DL.Movey2,
                    DL.PaintColor,
                    GI.minX,
                    GI.maxX,
                    GI.minY,
                    GI.maxY,
                    GI.c,
                    GI._cax,
                    GI.lemac,
                    GI.refsta,
                    GI.k,
                    ROW_NUMBER() OVER (PARTITION BY DL.NameLine ORDER BY DL.CreatedAt DESC, GI.CreatedAt DESC) AS RowNum
                FROM DataLine AS DL
                LEFT JOIN Graphics_item AS GI 
                    ON DL.AircraftType = GI.AircraftType 
                    AND DL.flight_bort = GI.flight_bort
                WHERE DL.AircraftType = @AircraftType AND DL.flight_bort = @flight_bort AND DL.NameLine IN ('MTOW', 'MLW', 'MZFW')
            )
            SELECT *
            FROM RankedData
            WHERE RowNum = 1 
            ORDER BY NameLine;"

                Dim command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@AircraftType", Form1.TextBox1.Text)
                command.Parameters.AddWithValue("@flight_bort", Form1.TextBox4.Text)

                Using reader As SqlDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim nameLine As String = reader("NameLine").ToString()

                        Dim pointsX0 = reader("PointsX0").ToString().Split(";"c)
                        Dim pointsY0 = reader("PointsY0").ToString().Split(";"c)
                        Dim pointsX1 = reader("PointsX1").ToString().Split(";"c)
                        Dim pointsY1 = reader("PointsY1").ToString().Split(";"c)

                        For i = 0 To Math.Min(pointsX0.Length, pointsY0.Length) - 1
                            pointsX0(i) = pointsX0(i).Replace(".", ",")
                            pointsY0(i) = pointsY0(i).Replace(".", ",")
                            pointsX1(i) = pointsX1(i).Replace(".", ",")
                            pointsY1(i) = pointsY1(i).Replace(".", ",")
                        Next

                        minX = Convert.ToInt32(reader("minX").ToString.Replace(".", ",")) / 1000
                        maxX = Convert.ToInt32(reader("maxX").ToString.Replace(".", ",")) / 1000
                        minY = Convert.ToInt32(reader("minY").ToString.Replace(".", ",")) / 1000
                        maxY = Convert.ToInt32(reader("maxY").ToString.Replace(".", ",")) / 1000

                        c = Convert.ToInt32(reader("c").ToString)
                        _cax = Single.Parse(reader("_cax").ToString)
                        lemac = Single.Parse(reader("lemac").ToString)
                        refsta = Single.Parse(reader("refsta").ToString)
                        k = Convert.ToInt32(reader("k").ToString)

                        Dim movex0 = reader("Movex0").ToString().Split(";"c)
                        Dim movey0 = reader("Movey0").ToString().Split(";"c)
                        Dim movex1 = reader("Movex1").ToString().Split(";"c)
                        Dim movey1 = reader("Movey1").ToString().Split(";"c)
                        Dim movex2 = reader("Movex2").ToString().Split(";"c)
                        Dim movey2 = reader("Movey2").ToString().Split(";"c)

                        For i = 0 To Math.Min(movex0.Length, movey0.Length) - 1
                            movex0(i) = movex0(i).Replace(".", ",")
                            movey0(i) = movey0(i).Replace(".", ",")
                            movex1(i) = movex1(i).Replace(".", ",")
                            movey1(i) = movey1(i).Replace(".", ",")
                            movex2(i) = movex2(i).Replace(".", ",")
                            movey2(i) = movey2(i).Replace(".", ",")
                        Next

                        Dim color = SKColor.Parse(reader("PaintColor").ToString())

                        Select Case nameLine
                            Case "MTOW"
                                If pointsX0.Length > 0 Then
                                    For i = 0 To pointsX0.Length - 1
                                        If Not String.IsNullOrWhiteSpace(pointsX0(i)) AndAlso Not String.IsNullOrWhiteSpace(pointsY0(i)) AndAlso
                                   Not String.IsNullOrWhiteSpace(pointsX1(i)) AndAlso Not String.IsNullOrWhiteSpace(pointsY1(i)) Then
                                            MtowLine.Add(Tuple.Create(Single.Parse(pointsX0(i)), Single.Parse(pointsY0(i)), Single.Parse(pointsX1(i)), Single.Parse(pointsY1(i))))
                                        End If
                                    Next
                                End If
                                If movex0.Length > 0 Then
                                    For i = 0 To movex0.Length - 1
                                        If Not String.IsNullOrWhiteSpace(movex0(i)) AndAlso Not String.IsNullOrWhiteSpace(movey0(i)) AndAlso
                                   Not String.IsNullOrWhiteSpace(movex1(i)) AndAlso Not String.IsNullOrWhiteSpace(movey1(i)) AndAlso
                                   Not String.IsNullOrWhiteSpace(movex2(i)) AndAlso Not String.IsNullOrWhiteSpace(movey2(i)) Then
                                            MtowCurve.Add(Tuple.Create(Single.Parse(movex0(i)), Single.Parse(movey0(i)), Single.Parse(movex1(i)), Single.Parse(movey1(i)), Single.Parse(movex2(i)), Single.Parse(movey2(i))))
                                        End If
                                    Next
                                End If
                                colormtow = color

                            Case "MLW"
                                If pointsX0.Length > 0 Then
                                    For i = 0 To pointsX0.Length - 1
                                        If Not String.IsNullOrWhiteSpace(pointsX0(i)) AndAlso Not String.IsNullOrWhiteSpace(pointsY0(i)) AndAlso
                                   Not String.IsNullOrWhiteSpace(pointsX1(i)) AndAlso Not String.IsNullOrWhiteSpace(pointsY1(i)) Then
                                            MlwLine.Add(Tuple.Create(Single.Parse(pointsX0(i)), Single.Parse(pointsY0(i)), Single.Parse(pointsX1(i)), Single.Parse(pointsY1(i))))
                                        End If
                                    Next
                                End If
                                If movex0.Length > 0 Then
                                    For i = 0 To movex0.Length - 1
                                        If Not String.IsNullOrWhiteSpace(movex0(i)) AndAlso Not String.IsNullOrWhiteSpace(movey0(i)) AndAlso
                                   Not String.IsNullOrWhiteSpace(movex1(i)) AndAlso Not String.IsNullOrWhiteSpace(movey1(i)) AndAlso
                                   Not String.IsNullOrWhiteSpace(movex2(i)) AndAlso Not String.IsNullOrWhiteSpace(movey2(i)) Then
                                            MlwCurve.Add(Tuple.Create(Single.Parse(movex0(i)), Single.Parse(movey0(i)), Single.Parse(movex1(i)), Single.Parse(movey1(i)), Single.Parse(movex2(i)), Single.Parse(movey2(i))))
                                        End If
                                    Next
                                End If
                                colormlw = color

                            Case "MZFW"
                                If pointsX0.Length > 0 Then
                                    For i = 0 To pointsX0.Length - 1
                                        If Not String.IsNullOrWhiteSpace(pointsX0(i)) AndAlso Not String.IsNullOrWhiteSpace(pointsY0(i)) AndAlso
                                   Not String.IsNullOrWhiteSpace(pointsX1(i)) AndAlso Not String.IsNullOrWhiteSpace(pointsY1(i)) Then
                                            MzfwLine.Add(Tuple.Create(Single.Parse(pointsX0(i)), Single.Parse(pointsY0(i)), Single.Parse(pointsX1(i)), Single.Parse(pointsY1(i))))
                                        End If
                                    Next
                                End If
                                If movex0.Length > 0 Then
                                    For i = 0 To movex0.Length - 1
                                        If Not String.IsNullOrWhiteSpace(movex0(i)) AndAlso Not String.IsNullOrWhiteSpace(movey0(i)) AndAlso
                                   Not String.IsNullOrWhiteSpace(movex1(i)) AndAlso Not String.IsNullOrWhiteSpace(movey1(i)) AndAlso
                                   Not String.IsNullOrWhiteSpace(movex2(i)) AndAlso Not String.IsNullOrWhiteSpace(movey2(i)) Then
                                            MzfwCurve.Add(Tuple.Create(Single.Parse(movex0(i)), Single.Parse(movey0(i)), Single.Parse(movex1(i)), Single.Parse(movey1(i)), Single.Parse(movex2(i)), Single.Parse(movey2(i))))
                                        End If
                                    Next
                                End If
                                colorzfw = color
                        End Select
                    End While
                End Using
            End Using

            mtow = New MTOW(Nothing, SkControl1.Width, SkControl1.Height, minX, maxX, minY, maxY, colormtow)
            mlw = New MLW(Nothing, SkControl1.Width, SkControl1.Height, minX, maxX, minY, maxY, colormlw)
            mzfw = New MZFW(Nothing, SkControl1.Width, SkControl1.Height, minX, maxX, minY, maxY, colorzfw)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


End Class