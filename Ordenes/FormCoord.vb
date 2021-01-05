Imports System.IO
Imports System.Configuration
Imports Negocio
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports System.Security.Policy
Imports System.Threading


Public Class FormCoord
    Protected WobjEstados As New Estados()
    Protected WobjOrden As New Orden()
    Protected WobjServicios As New Servicios()
    Protected WobjUnidades As New Unidades()
    Protected ListaOtDelete = New List(Of String)
    Protected WobjOrdenDetalle As New OrdenDetalle()
    Private Sub FormBuscarOT_Load(ByVal sender As System.Object, ByVal e As EventArgs) Handles MyBase.Load
        Try
            LlenaGrilla()
            LlenaEstados()
            LlenaServicios()
            LlenaUnidad()
        Catch ex As Exception
            MessageBox.Show("Error ocurrido al buscar información", "Cargar Grilla", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

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
    Sub LlenaUnidad()

        Dim unidad As DataSet = WobjUnidades.QryGetAllUnidadDes()
        Dim newRow As DataRow = unidad.Tables(0).NewRow
        newRow("IdUnidad") = "-1"
        newRow("Nombre") = ""

        unidad.Tables(0).Rows.Add(newRow)

        'txtUdestino.DataSource = unidad.Tables(0)
        'txtUdestino.DisplayMember = "Nombre"
        'txtUdestino.ValueMember = "IdUnidad"

        'txtUdestino.SelectedValue = -1

    End Sub

    Private Sub btnBuscar_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles btnBuscar.Click
        lblNumero.Text = String.Empty
        lblServicio.Text = String.Empty
        lblDestino.Text = String.Empty
        lblEstado.Text = String.Empty
        Try
            GeneraFiltro()
            DataGridView1.DataSource = WobjOrden.GetOrdenesByFiltro().Tables(0)
        Catch ex As Exception
            MessageBox.Show("Ocurrio un error al cargar las OT", "Buscando OT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

        ''If (txtUdestino.Text = "") Then
        'WobjOrden.UnidadDestino = 0
        'Else
        'WobjOrden.UnidadDestino = txtUdestino.SelectedValue
        'End If

        If (txtEstado.Text = "") Then
            WobjOrden.Estado = 0
        Else
            WobjOrden.Estado = txtEstado.SelectedValue
        End If
    End Sub

    Private Sub DataGridView1_RowContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)
        Dim rowIndex = e.RowIndex
        If (rowIndex <> -1) Then
            Dim row As DataGridViewRow = DataGridView1.Rows(rowIndex)
            lblNumero.Text = RTrim(row.Cells(0).Value)
            lblServicio.Text = RTrim(row.Cells(4).Value)
            lblDestino.Text = RTrim(row.Cells(9).Value)
            lblEstado.Text = RTrim(row.Cells(2).Value)
            txtSolicitante.Text = RTrim(row.Cells(13).Value)
            txtPiso.Text = IIf(RTrim(row.Cells(8).Value) = 0, "", RTrim(row.Cells(8).Value))
            txtModulo.Text = IIf(RTrim(row.Cells(7).Value) = "", "", RTrim(row.Cells(7).Value))
            txtRecinto.Text = RTrim(row.Cells(6).Value)

        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles imprimir.Click
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
                'ListaOtDelete.Add(fileName)


                'Thread.Sleep(3000)
                'If (File.Exists(fileName)) Then
                '    File.Delete(fileName)
                'End If

            End If

        Catch ex As Exception
            MessageBox.Show("Ocurrio un error al intentar Generar el PDF", "Generar PDF", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub


    Private Sub FormBuscarOT_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed

        For i As Integer = 0 To ListaOtDelete.count() - 1
            If (File.Exists(ListaOtDelete(i))) Then
                'For Each clsProcess In From clsProcess1 In Process.GetProcesses() Where (clsProcess1.ProcessName.Contains("AcroRd32"))
                '    clsProcess.Kill()
                'Next
                File.Delete(ListaOtDelete(i))
            End If
        Next
    End Sub


    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim rowIndex = e.RowIndex
        If (rowIndex <> -1) Then
            Dim row As DataGridViewRow = DataGridView1.Rows(rowIndex)
            DataGridView1.Rows(DataGridView1.CurrentRow.Index).Selected = True
            lblNumero.Text = RTrim(row.Cells(0).Value)
            lblServicio.Text = RTrim(row.Cells(4).Value)
            lblDestino.Text = RTrim(row.Cells(9).Value)
            lblEstado.Text = RTrim(row.Cells(2).Value)
            txtSolicitante.Text = RTrim(row.Cells(13).Value)
            txtPiso.Text = IIf(RTrim(row.Cells(8).Value) = "", "", RTrim(row.Cells(8).Value))
            txtModulo.Text = IIf(RTrim(row.Cells(7).Value) = "", "", RTrim(row.Cells(7).Value))
            txtRecinto.Text = RTrim(row.Cells(6).Value)

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

        End If
    End Sub

    Private Sub txtNumero_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNumero.TextChanged

    End Sub

    Private Sub txtSolucionProblema_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSolucionProblema.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class