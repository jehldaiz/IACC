Imports System.IO
Imports Negocio
Imports System.Reflection

Public Class Login
    Protected Encrypt As New Encrypt()
    Protected WobjUser As New Usuario()
    Public WobjUserSession As New Usuario()
    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Try
            WobjUser.UserName = txtuserName.Text
            WobjUser.PassUsuario = Encrypt.Encrypt_QueryString(txtPass.Text)
            If (WobjUser.GetUsuario().Tables(0).Rows.Count = 0) Then
                MessageBox.Show("El nombre de Usuario incorrecto o no Existe", "Ingreso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            Else
                For Each o In WobjUser.GetUsuario().Tables(0).Rows
                    WobjUserSession.UserName = o("UserName")
                    WobjUserSession.NomCompleto = o("NomCompleto")
                    If IsDBNull(o("RutUsuario")) Then
                        WobjUserSession.RutUsuario = 0
                    Else
                        WobjUserSession.RutUsuario = o("RutUsuario")
                    End If

                    If IsDBNull(o("RutDvUsuario")) Then
                        WobjUserSession.RutDvUsuario = ""
                    Else
                        WobjUserSession.RutDvUsuario = o("RutDvUsuario")
                    End If

                    If IsDBNull(o("MailUsuario")) Then
                        WobjUserSession.MailUsuario = ""
                    Else
                        WobjUserSession.MailUsuario = o("MailUsuario")
                    End If

                    WobjUserSession.PassUsuario = o("PassUsuario")
                    WobjUserSession.PassDefault = o("PassDefault")
                    WobjUserSession.FechaCreacion = o("FechaCreacion")
                    WobjUserSession.Activo = o("Activo")
                    WobjUserSession.IdPerfil = o("IdPerfil")
                    WobjUserSession.NombrePerfil = o("NombrePerfil")
                    WobjUserSession.IdUnidada = o("IdUnidad")
                    WobjUserSession.NombreUnidad = o("NombreUnidad")
                Next
                If (WobjUserSession.PassUsuario <> Encrypt.Encrypt_QueryString(txtPass.Text)) Then
                    MessageBox.Show("La contraseña es Incorrecta", "Ingreso", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    Hide()
                    FormPrincipal.Show()
                End If


            End If
        Catch ex As Exception
            MessageBox.Show("Ocurrio un error al intentar ingresar", "Ingreso", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub

    Private Sub txtuserName_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtuserName.KeyPress

        If Asc(e.KeyChar) = 13 Then
            If txtuserName.Text <> "" Then
                txtPass.Focus()
                e.Handled = True
            End If
        End If

    End Sub

    Private Sub txtPass_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtPass.KeyPress
        If Asc(e.KeyChar) = 13 Then

            If txtPass.Text <> "" Then
                OK.Enabled = True

                OK.Focus()
                e.Handled = True

            End If
            
        End If
    End Sub
End Class
