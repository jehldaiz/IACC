Imports System.IO
Imports Negocio
Imports iTextSharp.text
Imports iTextSharp.text.pdf

Public Class FormOTtecnico
    Protected WobjEstados As New Estados()
    Protected WobjOrden As New Orden()
    Protected WobjOrdenDetalle As New OrdenDetalle()
    Protected WobjServicios As New Servicios()
    Protected WobjUnidades As New Unidades()
    Protected WobjMateriales As New Materiales()
    Dim _listaMaterialesOt
    Private Sub FormOTtecnico_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            EnabledPanelContents(Panel8, False)
            EnabledPanelContents(Panel4, False)
            LlenaGrilla()
            LlenaEstados()
            LlenaServicios()
        Catch ex As Exception
            MessageBox.Show("Error ocurrido al buscar información", "Cargar Grilla", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub EnabledPanelContents(panel As Panel, enabled As Boolean)
        For Each item As Control In panel.Controls
            item.Enabled = enabled
        Next
    End Sub

    Sub LlenaGrilla()
        WobjOrden.UnidadDestino = CInt(Login.WobjUserSession.IdUnidada)
        DataGridView1.DataSource = WobjOrden.QryGetAllOrdenesByUnidad().Tables(0)
        DataGridView1.Columns("IdEstado").Visible = False
        DataGridView1.Columns("IdServicio").Visible = False
        DataGridView1.Columns("IdUnidad").Visible = False
        DataGridView1.Columns("IdCargoSol").Visible = False
        DataGridView1.Columns("RECINTO").Visible = False
        DataGridView1.Columns("MODULO").Visible = False
        DataGridView1.Columns("PISO").Visible = False
        DataGridView1.Columns("SUPERVISOR").Visible = False
    End Sub
    Sub LlenaEstados()

        Dim estados As DataSet = WobjEstados.QryGetAllEstados()
        Dim newRow As DataRow = estados.Tables(0).NewRow
        newRow("IdEstado") = "-1"
        newRow("NomEstado") = ""

        estados.Tables(0).Rows.Add(newRow)

        txtEstado.DataSource = estados.Tables(0)
        txtEstado.DisplayMember = "NomEstado"
        txtEstado.ValueMember = "IdEstado"

        txtEstado.SelectedValue = -1

    End Sub
    Sub LlenaServicios()

        Dim servicios As DataSet = WobjServicios.QryGetAllServicios()
        Dim newRow As DataRow = servicios.Tables(0).NewRow
        newRow("IdServicio") = "-1"
        newRow("NomServicio") = ""

        servicios.Tables(0).Rows.Add(newRow)

        txtServicio.DataSource = servicios.Tables(0)
        txtServicio.DisplayMember = "NomServicio"
        txtServicio.ValueMember = "IdServicio"

        txtServicio.SelectedValue = -1

    End Sub
    Private Sub GeneraFiltro()

        If (txtNumero.Text = "") Then
            WobjOrden.NOrden = 0
        Else
            WobjOrden.NOrden = CInt(txtNumero.Text)
        End If

        If (txtServicio.Text = "") Then
            WobjOrden.Servicio = 0
        Else
            WobjOrden.Servicio = txtServicio.SelectedValue
        End If

        WobjOrden.UnidadDestino = CInt(Login.WobjUserSession.IdUnidada)


        If (txtEstado.Text = "") Then
            WobjOrden.Estado = 0
        Else
            WobjOrden.Estado = txtEstado.SelectedValue
        End If
    End Sub
    Private Sub btnBuscar_Click(sender As System.Object, e As System.EventArgs) Handles btnBuscar.Click
        Try
            GeneraFiltro()
            DataGridView1.DataSource = WobjOrden.GetOrdenesByFiltro().Tables(0)
        Catch ex As Exception
            MessageBox.Show("Ocurrio un error al cargar las OT", "Buscando OT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub imprimir_Click(sender As System.Object, e As System.EventArgs) Handles imprimir.Click
        Try
            If (lblNumero.Text = "") Then
                MessageBox.Show("Debe seleccionar una OT", "Generar PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                Dim fileName As String = Path.GetTempPath() + Guid.NewGuid().ToString() + "-OT-" + lblNumero.Text.ToString() + ".pdf"
                Dim documento As New Document(PageSize.A4, 50, 50, 25, 25)
                PdfWriter.GetInstance(documento, New FileStream(fileName, FileMode.Create))
                documento.Open()

                Dim wvarPdf As New GenerarPdf()
                wvarPdf.CreaPdf(documento, CInt(lblNumero.Text))
                documento.Close()

                Dim prc As Process = New Process()
                prc.StartInfo.FileName = fileName
                prc.Start()
            End If

        Catch ex As Exception
            MessageBox.Show("Ocurrio un error al intentar Generar el PDF", "Generar PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnAddMaterial_Click(sender As System.Object, e As System.EventArgs) Handles btnAddMaterial.Click
        Dim M = New Materiales()
        If (lblNumero.Text = "") Then
            MessageBox.Show("Debe Seleccionar una OT", "Ingreso de Materiales", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If (txtNomItem.Text = "" Or txtCantidad.Text = "") Then
                MessageBox.Show("Debe ingresar el nombre y cantidad", "Ingreso de Materiales", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else

                M.NOrden = CInt(lblNumero.Text)
                M.MaterialNom = txtNomItem.Text
                M.Cantidad = CInt(txtCantidad.Text)
                M.Precio = 0
                M.Id = 0
                _listaMaterialesOt.add(M)
                LlenarMateriales()
                LimpiarCamposmateriales()

            End If
        End If
    End Sub

    Sub LlenarMateriales()
        DataGridView2.Rows.Clear()
        For i As Integer = 0 To _listaMaterialesOt.Count - 1
            DataGridView2.Rows.Add()
            DataGridView2.Rows(i).Cells(0).Value = _listaMaterialesOt(i).NOrden
            DataGridView2.Rows(i).Cells(1).Value = _listaMaterialesOt(i).MaterialNom
            DataGridView2.Rows(i).Cells(2).Value = _listaMaterialesOt(i).Cantidad
            DataGridView2.Rows(i).Cells(3).Value = _listaMaterialesOt(i).Precio
            DataGridView2.Rows(i).Cells(4).Value = _listaMaterialesOt(i).Id
        Next

    End Sub
    Sub LlenarGrillaMateriales()
        Try
            Dim TB As DataSet = WobjMateriales.GetMaterialesByOt()
            If TB.Tables(0).Rows.Count > 0 Then
                DataGridView2.Rows.Clear()
                For i As Integer = 0 To TB.Tables(0).Rows.Count - 1
                    Dim M As New Materiales()
                    DataGridView2.Rows.Add()
                    DataGridView2.Rows(i).Cells(0).Value = TB.Tables(0).Rows(i).Item("OT")
                    DataGridView2.Rows(i).Cells(1).Value = TB.Tables(0).Rows(i).Item("NomItem")
                    DataGridView2.Rows(i).Cells(2).Value = TB.Tables(0).Rows(i).Item("CantidadItem")
                    DataGridView2.Rows(i).Cells(3).Value = TB.Tables(0).Rows(i).Item("ValorItem")
                    DataGridView2.Rows(i).Cells(4).Value = TB.Tables(0).Rows(i).Item("id")
                    M.NOrden = CInt(TB.Tables(0).Rows(i).Item("OT"))
                    M.MaterialNom = TB.Tables(0).Rows(i).Item("NomItem")
                    M.Cantidad = CInt(TB.Tables(0).Rows(i).Item("CantidadItem"))
                    M.Precio = CInt(TB.Tables(0).Rows(i).Item("ValorItem"))
                    M.Id = CInt(TB.Tables(0).Rows(i).Item("id"))
                    _listaMaterialesOt.add(M)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Sub LimpiarCamposmateriales()
        txtNomItem.Text = ""
        txtCantidad.Text = ""
    End Sub

    Private Sub btnEliminarMaterial_Click(sender As System.Object, e As System.EventArgs) Handles btnEliminarMaterial.Click
        If (DataGridView2.Rows.Count > 0) Then
            If (DataGridView2.Rows(DataGridView2.CurrentRow.Index).Selected = True) Then
                If (Convert.ToString(DataGridView2.CurrentRow.Cells(0).Value) <> "") Then
                    If (Convert.ToString(DataGridView2.CurrentRow.Cells(4).Value) = "0") Then
                        _listaMaterialesOt.RemoveAt(DataGridView2.CurrentRow.Index)
                        DataGridView2.Rows.RemoveAt(DataGridView2.CurrentRow.Index)
                    Else
                        WobjMateriales.NOrden = CInt(lblNumero.Text)
                        WobjMateriales.Id = CInt(DataGridView2.CurrentRow.Cells(4).Value)
                        If (MessageBox.Show("¿Está Seguro que Desea Eliminar el Item: " + DataGridView2.CurrentRow.Cells(1).Value, "Sistema Ordenes de Trabajo.", MessageBoxButtons.YesNo, MessageBoxIcon.Error) = DialogResult.Yes) Then
                            For Each r In WobjMateriales.Delete_ItemOt().Tables(0).Rows
                                If (r("RESULT") = "0") Then
                                    MessageBox.Show("Item Eliminado Correctamente.", "Ingreso de Materiales.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Else
                                    MessageBox.Show("Ocurrio un error al eliminar el Item.", "Ingreso de Materiales.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                End If
                            Next
                            LlenarGrillaMateriales()
                        End If
                    End If
                Else
                    MessageBox.Show("No Existe Ningun Elemento en la Lista.", "Ingreso de Materiales.", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    DataGridView2.ClearSelection()
                End If
            Else
                MessageBox.Show("Por Favor Seleccione Item a Eliminar de la Lista.", "Ingreso de Materiales.", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        Else
            MessageBox.Show("No Existe Ningun Elemento en la Lista", "Ingreso de Materiales.", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If (DataGridView2.Rows.Count > 0) Then
            DataGridView2.Rows(DataGridView2.CurrentRow.Index).Selected = True
        End If
    End Sub

    Private Sub DataGridView2_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If (DataGridView2.Rows.Count > 0) Then
            DataGridView2.Rows(DataGridView2.CurrentRow.Index).Selected = True
        End If
    End Sub

    Private Sub txtCantidad_KeyPress(sender As System.Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtCantidad.KeyPress
        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
    Private Function Guardar() As String
        Dim NomJefeUnidad
        Dim msg = String.Empty

        Dim ordenDetalle As DataSet = WobjUnidades.GetJefeUnidad(Login.WobjUserSession.IdUnidada)
        If ordenDetalle.Tables(0).Rows.Count > 0 Then
            For Each d In ordenDetalle.Tables(0).Rows
                NomJefeUnidad = d("NomCompleto")
            Next
            WobjOrdenDetalle.NOrden = CInt(lblNumero.Text)
            WobjOrdenDetalle.Fecha = DateAndTime.Now
            WobjOrdenDetalle.TecnicoUnidad = Login.WobjUserSession.NomCompleto
            WobjOrdenDetalle.JefeUnidad = NomJefeUnidad
            WobjOrdenDetalle.SolucionProblema = txtSolucionProblema.Text
            WobjOrdenDetalle.EstadoId = CInt(cmbEstadoCambio.SelectedValue)
            WobjOrdenDetalle.Usuario = Login.WobjUserSession.UserName
            WobjOrdenDetalle.Comentario = txtComentario.Text
            WobjOrdenDetalle.JefeServicio = txtJefeServicio.Text.ToUpper()

            For Each o In WobjOrdenDetalle.INS_DetalleOrden().Tables(0).Rows
                If (o("RESULT") = "1") Then
                    msg = "OK"
                End If
            Next

            If (msg <> String.Empty) Then
                Return msg
            End If
        Else
            msg = "No se puede Guardar la información ya que no existe un Jefe de Unidad"
        End If
        Return msg
    End Function
    Private Function GuardarMateriales() As String
        Dim msg = String.Empty
        Dim b As Integer
        For Each row As DataGridViewRow In Me.DataGridView2.Rows
            WobjMateriales.NOrden = CInt(row.Cells(0).Value)
            WobjMateriales.MaterialNom = row.Cells(1).Value.ToString()
            WobjMateriales.Cantidad = CInt(row.Cells(2).Value)
            WobjMateriales.Precio = CInt(row.Cells(3).Value)
            WobjMateriales.Id = CInt(row.Cells(4).Value)

            For Each e In WobjMateriales.INS_ItemOt().Tables(0).Rows
                If (e("RESULT") <> "1") Then
                    msg = msg + vbNewLine + "El Item: " + row.Cells(1).Value.ToString() + "No pudo ser ingresado"
                Else
                    msg = "OK"
                End If
            Next
        Next
        Return msg
    End Function

    Private Function LimpiarGuardarOk()
        EnabledPanelContents(Panel8, False)
        EnabledPanelContents(Panel4, False)
        txtComentario.Text = String.Empty
        txtSolucionProblema.Text = String.Empty
        lblNumero.Text = String.Empty
        lblServicio.Text = String.Empty
        lblDestino.Text = String.Empty
        lblEstado.Text = String.Empty
        txtSolicitante.Text = String.Empty
        txtPiso.Text = String.Empty
        txtModulo.Text = String.Empty
        txtRecinto.Text = String.Empty
        lblIdServicio.Text = String.Empty
        lblIdDestino.Text = String.Empty
        lblIdEstado.Text = String.Empty
        DataGridView2.Rows.Clear()
        lblOtSelec.Text = String.Empty
    End Function

    Private Sub btnIngresar_Click(sender As System.Object, e As System.EventArgs) Handles btnIngresar.Click
        Dim mensage As String
        Try
            If (lblIdEstado.Text = "6" Or lblIdEstado.Text = "2") Then
                If (cmbEstadoCambio.Text = String.Empty) Then
                    MessageBox.Show("Seleccione el Nuevo estado de la OT", "Sistema Ordenes de Trabajo.", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                Else
                    mensage = Guardar()
                    LlenaGrilla()
                End If
            End If

            If (lblIdEstado.Text = "1" And cmbEstadoCambio.Text <> "") Then
                If (cmbEstadoCambio.SelectedValue = 2 And txtComentario.Text = String.Empty) Then
                    MessageBox.Show("Si deja la OT Pendiente debe ingresar un Comentario", "Sistema Ordenes de Trabajo.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                Else
                    If (cmbEstadoCambio.SelectedValue = "3" And (lblIdEstado.Text = "1" Or lblIdEstado.Text = "2")) Then
                        If (txtSolucionProblema.Text = String.Empty) Then
                            MessageBox.Show("Si finaliza la OT debe ingresar la Solución del Problema", "Sistema Ordenes de Trabajo.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Exit Sub
                        End If
                        If (txtJefeServicio.Text = String.Empty) Then
                            MessageBox.Show("Si finaliza la OT debe ingresar el Supervisio o Jefe de Servicio", "Sistema Ordenes de Trabajo.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Exit Sub
                        End If
                    End If
                    mensage = Guardar()
                    LlenaGrilla()
                End If

            End If
            If (lblIdEstado.Text = "1") Then
                If (DataGridView2.Rows.Count > 0) Then
                    If (MessageBox.Show("¿Está Seguro que Desea Guardar los Mateirales en la OT: " + lblNumero.Text, "Sistema Ordenes de Trabajo.", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes) Then
                        mensage = GuardarMateriales()
                        WobjMateriales.NOrden = CInt(lblNumero.Text)
                        LlenarGrillaMateriales()
                    End If

                End If
            End If

            If (mensage <> "") Then
                If (mensage = "OK") Then
                    MessageBox.Show("Registro Guardado Correctamente", "Sistema Ordenes de Trabajo.", MessageBoxButtons.OK, MessageBoxIcon.None)
                    LimpiarGuardarOk()
                Else
                    MessageBox.Show(mensage, "Sistema Ordenes de Trabajo.", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

            End If
        Catch ex As Exception
            MessageBox.Show("Ocurrio un error al Guardar la Información", "Sistema Ordenes de Trabajo", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmbEstadoCambio_SelectedValueChanged(sender As System.Object, e As System.EventArgs) Handles cmbEstadoCambio.SelectedValueChanged
        If cmbEstadoCambio.SelectedValue Is Nothing Then Exit Sub
        Try
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        DataGridView2.Rows.Clear()
        _listaMaterialesOt = New List(Of Materiales)
        Dim rowIndex = e.RowIndex
        If (rowIndex <> -1) Then
            DataGridView1.Rows(DataGridView1.CurrentRow.Index).Selected = True
            Dim row As DataGridViewRow = DataGridView1.Rows(rowIndex)
            lblNumero.Text = RTrim(row.Cells(0).Value)
            lblServicio.Text = RTrim(row.Cells(4).Value)
            lblDestino.Text = RTrim(row.Cells(9).Value)
            lblEstado.Text = RTrim(row.Cells(2).Value)
            txtSolicitante.Text = RTrim(row.Cells(13).Value)
            txtPiso.Text = IIf(RTrim(row.Cells(8).Value) = "", "", RTrim(row.Cells(8).Value))
            txtModulo.Text = IIf(RTrim(row.Cells(7).Value) = "", "", RTrim(row.Cells(7).Value))
            txtRecinto.Text = RTrim(row.Cells(6).Value)
            lblIdServicio.Text = RTrim(row.Cells(5).Value)
            lblIdDestino.Text = RTrim(row.Cells(10).Value)
            lblIdEstado.Text = RTrim(row.Cells(3).Value)
            If Convert.ToString(RTrim(row.Cells(16).Value.ToString()) <> "") Then
                txtJefeServicio.Text = row.Cells(16).Value.ToString()
            Else
                txtJefeServicio.Text = ""
            End If

            lblOtSelec.Text = "OT SELECCIONADA : " + RTrim(row.Cells(0).Value)
            Dim estados As DataSet = WobjEstados.QryGetAllEstados()
            WobjOrdenDetalle.NOrden = Convert.ToInt32(lblNumero.Text)

            ' ******************************************************************************************
            Dim getTecnico As DataSet = WobjOrdenDetalle.GetTecnicoMaxRegistro()
            If (getTecnico.Tables(0).Rows.Count > 0) Then
                For Each o In getTecnico.Tables(0).Rows
                    txtComentario.Text = o("COMENTARIO").ToString()
                    txtSolucionProblema.Text = o("SOLUCION_PROBLEMA").ToString()
                Next
            End If
            ' ******************************************************************************************

            Dim newRow As DataRow = estados.Tables(0).NewRow
            newRow("IdEstado") = "-1"
            newRow("NomEstado") = ""

            Dim filteredTable As DataTable
            If (lblIdEstado.Text = "6" Or lblIdEstado.Text = "2") Then
                EnabledPanelContents(Panel8, False)
                EnabledPanelContents(Panel4, False)
                cmbEstadoCambio.Enabled = True
                btnIngresar.Enabled = True
                Label7.Enabled = True
                filteredTable = (From n In estados.Tables(0).AsEnumerable()
               Where n.Field(Of Int32)("IdEstado") = 1 Select n).CopyToDataTable()
            End If

            If (lblIdEstado.Text = "1") Then
                EnabledPanelContents(Panel8, False)
                EnabledPanelContents(Panel4, False)
                cmbEstadoCambio.Enabled = True
                EnabledPanelContents(Panel4, True)
                Label7.Enabled = True
                Label6.Enabled = True
                Label15.Enabled = True
                txtSolucionProblema.Enabled = True
                txtJefeServicio.Enabled = True
                Label18.Enabled = True
                Label17.Enabled = True
                txtComentario.Enabled = True
                filteredTable = (From n In estados.Tables(0).AsEnumerable()
               Where n.Field(Of Int32)("IdEstado") = 2 Or n.Field(Of Int32)("IdEstado") = 3 Select n).CopyToDataTable()
            End If

            If (lblIdEstado.Text = "3") Then
                EnabledPanelContents(Panel8, False)
                EnabledPanelContents(Panel4, False)
                cmbEstadoCambio.Enabled = False
                filteredTable = (From n In estados.Tables(0).AsEnumerable()
               Where n.Field(Of Int32)("IdEstado") = 2 Or n.Field(Of Int32)("IdEstado") = 3 Select n).CopyToDataTable()
            End If

            cmbEstadoCambio.DataSource = filteredTable
            cmbEstadoCambio.DisplayMember = "NomEstado"
            cmbEstadoCambio.ValueMember = "IdEstado"
            cmbEstadoCambio.SelectedValue = -1

            WobjMateriales.NOrden = CInt(RTrim(row.Cells(0).Value))
            LlenarGrillaMateriales()
            txtCantidad.Text = ""
            _txtNomItem.Text = ""
            DataGridView2.ClearSelection()
        End If
    End Sub
End Class