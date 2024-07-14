Imports System.Drawing.Printing

Public Class Form15

    Private MyImage As Bitmap
    Public Property Mtow As New PointF(0, 0)
    Public Property Mzfw As New PointF(0, 0)
    Public Property MtowWeight As Integer = 0
    Public Property MzfwWeight As Integer = 0
    Public Property MtowIndex As Single = 0
    Public Property MzfwIndex As Single = 0

    Public Sub New(Mtow As PointF, Mzfw As PointF, MtowWeight As Integer, MzfwWeight As Integer, MtowIndex As Single, MzfwIndex As Single)

        ' Этот вызов является обязательным для конструктора.
        InitializeComponent()

        ' Добавить код инициализации после вызова InitializeComponent().
        Me.Mtow = Mtow
        Me.Mzfw = Mzfw
        Me.MzfwWeight = MzfwWeight
        Me.MtowWeight = MtowWeight
        Me.MtowIndex = MtowIndex
        Me.MzfwIndex = MzfwIndex
    End Sub


    Public Sub ShowMyImage(imageToDisplay As Bitmap)
        If PictureBox1.Image IsNot Nothing Then
            PictureBox1.Image.Dispose()
            PictureBox1.Image = Nothing
        End If

        If MyImage IsNot Nothing Then
            MyImage.Dispose()
            MyImage = Nothing
        End If


        MyImage = New Bitmap(800, 450)

        Using g As Graphics = Graphics.FromImage(MyImage)
            g.DrawImage(imageToDisplay, 0, 0, 800, 450)
        End Using


        imageToDisplay.Dispose()

        PictureBox1.Image = MyImage
        PictureBox1.Location = New Point(0, 0) ' or any other desired location
        PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
        PictureBox1.Width = 800
        PictureBox1.Height = 450
    End Sub
    Private Function Interpolate(value As Single, fromMin As Single, fromMax As Single, toMin As Single, toMax As Single) As Single
        Return (value - fromMin) * (toMax - toMin) / (fromMax - fromMin) + toMin
        '(x-x1)/(x2-x1)=(y-y1)/(y2-y1)
    End Function

    ' Метод для преобразования значений веса и индекса в координаты пикселей
    Private Function MapToPixelCoordinates(weight As Integer, index As Single) As PointF

        ' Преобразование веса в координату X

        Dim x As Single = Interpolate(index, MzfwIndex, MtowIndex, Mzfw.X, Mtow.X)

        ' Преобразование веса в координату Y
        Dim y As Single = Interpolate(weight, MzfwWeight, MtowWeight, Mzfw.Y, Mtow.Y)

        Return New PointF(x, y)
    End Function

    ' Метод для рисования точки на графике
    Private Sub PlotPoint(weight As Integer, index As Single, CAX As Single, LableCAX As String)

        Dim point As PointF = MapToPixelCoordinates(weight, index)
        ' Рисование точки на PictureBox
        ' g.FillEllipse(Brushes.Red, point.X - 4, point.Y - 4, 8, 8)
        ' Вывод значений координат
        'g.DrawString($"({LableCAX}: {CAX}%)", New Font("Arial", 8), Brushes.Black, point.X + 4, point.Y - 4)

        If MyImage Is Nothing Then Exit Sub

        Try

            Using g As Graphics = Graphics.FromImage(MyImage)


                g.FillEllipse(Brushes.Red, point.X - 4, point.Y - 4, 8, 8)
                ' Вывод значений координат
                g.DrawString($"({LableCAX}: {CAX}%)", New Font("Arial", 8), Brushes.Black, point.X + 4, point.Y - 4)

            End Using

            PictureBox1.Refresh()
        Catch ex As Exception
            MsgBox("Ошибка вывода изображения " & ex.Message, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Обработчик события для кнопки
    Public Sub Exempl(LableCAX As String, CAX As Single, Index As Single, Weight As Integer)
        ' Пример добавления точек на график с различными значениями веса и индекса
        ' Using g As Graphics = PictureBox1.CreateGraphics()
        'PlotPoint(g, 40000, 60) ' Примbер веса 55000 и индекс 60
        PlotPoint(Weight, Index, CAX, LableCAX)
        'End Using
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Beep()
        End Try
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Try
            ' Проверяем, что PictureBox содержит изображение
            If PictureBox1.Image IsNot Nothing Then
                ' Определяем размер изображения и области для печати
                Dim img As Image = PictureBox1.Image
                Dim printArea As Rectangle = e.MarginBounds

                ' Сохраняем пропорции изображения
                If img.Width / img.Height > printArea.Width / printArea.Height Then
                    printArea.Height = CInt(img.Height / img.Width * printArea.Width)
                Else
                    printArea.Width = CInt(img.Width / img.Height * printArea.Height)
                End If

                ' Печатаем изображение
                e.Graphics.DrawImage(img, printArea)
            Else
                ' Если изображение отсутствует, выводим сообщение
                e.Graphics.DrawString("Нет изображения для печати.", New Font("Times New Roman", 12), Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top)
            End If

            ' Устанавливаем флаг HasMorePages в False, т.к. печатаем только одну страницу
            e.HasMorePages = False
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Beep()
        End Try
    End Sub
End Class