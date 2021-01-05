Imports System.IO
Imports Negocio
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class IngresoOT

    Protected WobjOrden As New Orden()
    Protected Unidades As New Unidades()
    Protected WobjServicios As New Servicios()
    Protected WobjCarSol As New CargosSolicitante()
    Private Sub IngresoOT_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            lblOperador.Text = Login.WobjUserSession.NomCompleto.ToUpper()
            LlenaServicios()
            LlenaUnidad()
            LlenaCargoSolicitante()
        Catch ex As Exception
            MessageBox.Show("Error ocurrido al buscar información", "Cargar Grilla", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    Sub LlenaUnidad()

        Dim unidad As DataSet = Unidades.QryGetAllUnidadDes()
        Dim newRow As DataRow = unidad.Tables(0).NewRow
        newRow("IdUnidad") = "-1"
        newRow("Nombre") = ""

        unidad.Tables(0).Rows.Add(newRow)

        cmbDestino.DataSource = unidad.Tables(0)
        cmbDestino.DisplayMember = "Nombre"
        cmbDestino.ValueMember = "IdUnidad"

        cmbDestino.SelectedValue = -1

    End Sub
    Sub LlenaServicios()

        Dim servicios As DataSet = WobjServicios.QryGetAllServicios()
        Dim newRow As DataRow = servicios.Tables(0).NewRow
        newRow("IdServicio") = "-1"
        newRow("NomServicio") = ""

        servicios.Tables(0).Rows.Add(newRow)

        cmbServicio.DataSource = servicios.Tables(0)
        cmbServicio.DisplayMember = "NomServicio"
        cmbServicio.ValueMember = "IdServicio"

        cmbServicio.SelectedValue = -1

    End Sub
    Sub LlenaCargoSolicitante()

        Dim cargoSolicitante As DataSet = WobjCarSol.QryGetAllCargoSolicitante()
        Dim newRow As DataRow = cargoSolicitante.Tables(0).NewRow
        newRow("IdCargoSol") = "-1"
        newRow("NomCargoSol") = ""

        cargoSolicitante.Tables(0).Rows.Add(newRow)

        cmbCargoSol.DataSource = cargoSolicitante.Tables(0)
        cmbCargoSol.DisplayMember = "NomCargoSol"
        cmbCargoSol.ValueMember = "IdCargoSol"

        cmbCargoSol.SelectedValue = -1

    End Sub
    Private Sub Button1_Click(sender As System.Object, e As EventArgs) Handles btnIngresar.Click
        Dim msg As String = ValidaIngreso()
        Try
            If (msg <> String.Empty) Then
                MessageBox.Show(msg, "Ingreso OT", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            Else
                WobjOrden.Fecha = DateAndTime.Now
                WobjOrden.Operador = Login.WobjUserSession.UserName
                If IsNothing(cmbServicio.SelectedValue) Then
                    MessageBox.Show("Errorrrrrr", "Ingreso OT", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return

                End If
                WobjOrden.Servicio = cmbServicio.SelectedValue
                WobjOrden.Recinto = txtRecinto.Text
                WobjOrden.UnidadDestino = cmbDestino.SelectedValue
                WobjOrden.DescripcionProblema = txtProblema.Text
                WobjOrden.Solicitante = txtSolicitante.Text
                WobjOrden.CargoSolicitante = cmbCargoSol.SelectedValue
                If (txtModulo.Text = "") Then
                    WobjOrden.Modulo = ""
                Else
                    WobjOrden.Modulo = txtModulo.Text
                End If

                If (txtPiso.Text = "") Then
                    WobjOrden.Piso = ""
                Else
                    WobjOrden.Piso = txtPiso.Text
                End If

                Dim servicio As String

                For Each o In WobjOrden.INS_Ordenes().Tables(0).Rows
                    If (o("RESULT") = "1") Then
                        MessageBox.Show("Registro ingresado correctamente", "Ingreso OT", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        txtProblema.Text = String.Empty
                        txtRecinto.Text = String.Empty
                        txtSolicitante.Text = String.Empty
                        cmbServicio.SelectedValue = -1
                        cmbDestino.SelectedValue = -1
                        cmbCargoSol.SelectedValue = -1
                        txtModulo.Text = String.Empty
                        txtPiso.Text = String.Empty

                        If (CInt(o("OT")) > 0) Then
                            Try
                                Dim fileName As String = Path.GetTempPath() + Guid.NewGuid().ToString() + "OT-" + o("OT").ToString() + ".pdf"
                                Dim documento As New Document(PageSize.A4, 50, 50, 25, 25)
                                PdfWriter.GetInstance(documento, New FileStream(fileName, FileMode.Create))
                                documento.Open()

                                Dim wvarPdf As New GenerarPdf()
                                wvarPdf.CreaPdf(documento, CInt(o("OT")))
                                documento.Close()

                                Dim prc As Process = New Process()
                                prc.StartInfo.FileName = fileName
                                prc.Start()

                                'Thread.Sleep(3000)
                                'If (File.Exists(fileName)) Then
                                '    File.Delete(fileName)
                                'End If

                                Dim mail = New EnviarCorreo()
                                Dim para = ""
                                Dim cc = ""
                                WobjOrden.NOrden = CInt(o("OT"))
                                Dim orden As DataSet = WobjOrden.QryGetOrdenesById()
                                If orden.Tables(0).Rows.Count > 0 Then
                                    For Each c In orden.Tables(0).Rows
                                        Unidades.IdUnidad = CInt(c("IdUnidad"))
                                        servicio = c("SERVICIO")
                                    Next
                                End If
                                Dim unid As DataSet = Unidades.QryGetAllUnidadDesById()
                                For Each x In unid.Tables(0).Rows
                                    If (x("CorreoTaller").ToString() = "" And x("CorreoJefe").ToString() = "") Then
                                        MessageBox.Show("Los Correos no fueron enviados, ya que la Unidad no posee direcciones de cooreos asignada" + vbNewLine + "Por favor avisar de la creación de la nueva orden a la unidad correspondiente vía telefono", "Envío de Correo", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        Return
                                    End If
                                    If (x("CorreoTaller").ToString() = "") Then
                                        para = x("CorreoJefe").ToString()
                                    End If
                                    If (x("CorreoJefe").ToString() = "") Then
                                        para = x("CorreoTaller").ToString()
                                    End If
                                    If (x("CorreoTaller").ToString() <> "" And x("CorreoJefe").ToString() <> "") Then
                                        para = x("CorreoTaller").ToString() + "," + x("CorreoJefe").ToString()
                                    End If

                                Next

                                Dim body As String = String.Empty
                                Dim reader As String = My.Resources.ResourceManager.GetObject("MailTemplate")
                                body = reader
                                body = body.Replace("{Norden}", o("OT").ToString())
                                body = body.Replace("{Estado}", "Generada")
                                body = body.Replace("{Servi}", servicio)
                                body = body.Replace("{Reci}", WobjOrden.Recinto.ToString())
                                body = body.Replace("{Sol}", WobjOrden.Solicitante.ToString())
                                body = body.Replace("{Piso}", WobjOrden.Piso.ToString())
                                body = body.Replace("{Modulo}", WobjOrden.Modulo.ToString())
                                body = body.Replace("{Des}", WobjOrden.DescripcionProblema.ToString())

                                mail.SendMail(para, cc, body)


                            Catch ex As Exception
                                MessageBox.Show(ex, MessageBoxButtons.OK, MessageBoxIcon.Error)
                                MessageBox.Show("Ocurrio un error al generar la OT", "Generar OT", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End Try
                        End If
                    Else
                        MessageBox.Show("La OT no fue generada", "Ingreso OT", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Next
            End If
        Catch ex As Exception
            ' MessageBox.Show("Ocurrio un error al ingrear la OT", "Ingreso OT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function ValidaIngreso() As String

        Dim msg = String.Empty

        If (cmbServicio.Text = String.Empty) Then
            msg = msg + "Seleccione Servicio"
        End If
        If (txtRecinto.Text = String.Empty) Then
            msg = msg + vbNewLine + "Ingrese Recinto"
        End If
        If (cmbDestino.Text = String.Empty) Then
            msg = msg + vbNewLine + "Seleccione Destino"
        End If
        If (txtSolicitante.Text = String.Empty) Then
            msg = msg + vbNewLine + "Ingrese Solicitante"
        End If
        If (cmbCargoSol.Text = String.Empty) Then
            msg = msg + vbNewLine + "Seleccione Cargo Solicitante"
        End If
        'If (cmbModulo.Text = String.Empty) Then
        '    msg = msg + vbNewLine + "Seleccione Modulo"
        'End If
        'If (cmbPiso.Text = String.Empty) Then
        '    msg = msg + vbNewLine + "Seleccione Piso"
        'End If
        If (txtProblema.Text = String.Empty) Then
            msg = msg + vbNewLine + "Ingrese Descripcion del Problema"
        End If
        Return msg
    End Function

    Private Sub lblOperador_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblOperador.Click

    End Sub

    Private Sub cmbServicio_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbServicio.SelectedIndexChanged

    End Sub

    Private Sub txtSolicitante_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSolicitante.TextChanged

    End Sub

   

    Private Function WobjN_ORDEN() As Object
        Throw New NotImplementedException
    End Function

    Private Sub txtPiso_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPiso.TextChanged

    End Sub

    Private Sub txtModulo_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtModulo.TextChanged

    End Sub
End Class