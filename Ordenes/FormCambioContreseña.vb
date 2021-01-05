Imports System.IO
Imports Negocio

Public Class FormCambioContreseña
    Protected WobjUser As New Usuario()
    Private Shared Sub Cancel_Click(sender As System.Object, e As EventArgs) Handles Cancel.Click
        Application.Exit()
    End Sub

    Private Sub OK_Click(sender As System.Object, e As EventArgs) Handles OK.Click
        Dim msg As String = ValidaIngreso()
        Try
            If (msg <> String.Empty) Then
                MessageBox.Show(msg, "Cambio de Contraseña", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            Else
                WobjUser.UserName = Login.WobjUserSession.UserName
                WobjUser.PassUsuario = Encrypt.Encrypt_QueryString(txtPassNuev.Text)
                For Each o In WobjUser.UPD_PassUser().Tables(0).Rows
                    If (o("RESULT") = "1") Then
                        MessageBox.Show("Registro ingresado correctamente", "Ingreso OT", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Login.WobjUserSession.PassUsuario = Encrypt.Encrypt_QueryString(txtPassNuev.Text)
                        FormPrincipal.Show()
                    End If
                Next
                ' Dim TargetPath = "d:\files\" & Path.GetFileName(MyFile.PostedFile.FileName)
                '    'y por ultimo se envia el archivo a el servidor esto e lo que hace que el archivo se 
                '    ' envie al el servidor
                '    MyFile.PostedFile.SaveAs(TargetPath)
                'MyFile.PostedFile.(TargetPath)
            End If
        Catch ex As Exception
            MessageBox.Show("Ocurrio un error al intentar cambiar la contraseña", "Ingreso OT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Private Function ValidaIngreso() As String
        Dim str As String = String.Empty

        If (txtPassAct.Text = String.Empty) Then
            str = str + "Ingrese Nueva Contraseña"
        End If
        If (txtPassNuev.Text = String.Empty) Then
            str = str + vbNewLine + "Ingrese la nueva contraseña"
        End If

        If (Encrypt.Encrypt_QueryString(txtPassNuev.Text) = Login.WobjUserSession.PassUsuario) Then
            str = str + vbNewLine + "La nueva contraseña no puede ser igual a la actual "
        End If
        Return str
    End Function
End Class