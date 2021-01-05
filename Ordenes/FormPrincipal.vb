Imports System.Runtime.Remoting
Imports Negocio

Public Class FormPrincipal
    Protected WobjvalidaRut As New ValidadorRut()
    Private Sub FormPrincipal_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        tUserName.Text = Login.WobjUserSession.UserName
        tRol.Text = Login.WobjUserSession.NombrePerfil.ToString()
        tunidad.Text = Login.WobjUserSession.NombreUnidad
        If (Login.WobjUserSession.PassDefault = 1) Then
            Hide()
            FormCambioContreseña.Show()
        End If
        If (Login.WobjUserSession.IdPerfil <> 8 Or Login.WobjUserSession.NombrePerfil <> "OPERADORES") Then
            IngresarOTToolStripMenuItem.Enabled = False
        End If
        If (Login.WobjUserSession.IdUnidada = 6 Or Login.WobjUserSession.NombrePerfil = "TECNICO") Then
            IngresarOTToolStripMenuItem.Enabled = True
        End If
    End Sub

    Private Sub BuscarOTToolStripMenuItem_Click(sender As System.Object, e As EventArgs) Handles BuscarOTToolStripMenuItem.Click

        If (Login.WobjUserSession.IdPerfil = 8 Or Login.WobjUserSession.NombrePerfil = "OPERADORES") Then
            FormBuscarOT.Show()
        End If
        If (Login.WobjUserSession.IdPerfil = 7 Or Login.WobjUserSession.NombrePerfil = "TECNICOS") Then
            FormOTtecnico.Show()
        End If

        If (Login.WobjUserSession.IdPerfil = 6 Or Login.WobjUserSession.NombrePerfil = "JEFE") Then
            FormJefes.Show()
        End If

        If (Login.WobjUserSession.IdPerfil = 10 Or Login.WobjUserSession.NombrePerfil = "VISUALIZADOR OT") Then
            FormBuscarOT.Show()
        End If


        ''Comentario para ver la actualizacion en el repositorio GitHub
        ''If (Login.WobjUserSession.IdPerfil = 11 Or Login.WobjUserSession.NombrePerfil = "COORDINADOR") Then
        ''FormCoord.Show()
        ''End If


    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As System.Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub IngresarOTToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles IngresarOTToolStripMenuItem.Click
        If (Login.WobjUserSession.IdPerfil = 8 Or Login.WobjUserSession.NombrePerfil = "OPERADORES") Then
            IngresoOT.Show()
        End If
        If (Login.WobjUserSession.IdUnidada = 6 Or Login.WobjUserSession.NombrePerfil = "TECNICOS") Then
            IngresoOT.Show()
        End If


    End Sub

    Private Sub AyudaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AyudaToolStripMenuItem.Click

    End Sub

    Private Sub ManualDeUsuarioToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ManualDeUsuarioToolStripMenuItem.Click
        ManualUsuario.Show()
    End Sub

    Private Sub InfoSistemaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InfoSistemaToolStripMenuItem.Click
        InfoSis.Show()

    End Sub
End Class