Public Class Loader
    Private Sub Loader_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Loader()
    End Sub
    Public Sub Loader(row As Integer)
        ProgressBar1.Value = row
    End Sub
End Class