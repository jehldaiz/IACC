Public Class ManualUsuario

    Private Sub AxAcroPDF1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxAcroPDF1.Enter

    End Sub

    Private Sub ManualUsuario_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim fileName = "C:\Users\Zeus\Desktop\V27-07-2020\ManualUsuario1.pdf"
        AxAcroPDF1.LoadFile(fileName)
    End Sub
End Class