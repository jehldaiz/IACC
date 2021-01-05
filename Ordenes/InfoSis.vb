Public Class InfoSis

    Private Sub InfoSis_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim fileName = "C:\Users\Zeus\Desktop\V27-07-2020\ManualSoftware.pdf"
        AxAcroPDF1.LoadFile(fileName)
    End Sub

    Private Sub AxAcroPDF1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AxAcroPDF1.Enter

    End Sub
End Class