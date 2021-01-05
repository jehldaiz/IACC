Imports System.IO
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports Negocio
Imports System.Resources
Imports System.Reflection

Public Class GenerarPdf
    Protected WobjOrden As New Orden()
    Protected WobjOrdenMateriales As New Materiales()
    Protected WobjOrdenDetalle As New OrdenDetalle()
    Public Function CreaPdf(ByVal documento As Document, ByVal ot As Integer) As Document

        '***********************PARTE SUPERIOR******************************** 
        WobjOrden.NOrden = ot
        Dim ordenActual = New Orden()
        Dim listaDetalleOt = New List(Of OrdenDetalle)
        Dim listaMaterialesOt = New List(Of Materiales)
        Dim orden As DataSet = WobjOrden.QryGetOrdenesById()
        Dim Comtario As String
        If orden.Tables(0).Rows.Count > 0 Then
            For Each o In orden.Tables(0).Rows
                ordenActual.NOrden = o("ORDEN")
                ordenActual.Fecha = o("FECHA")
                ordenActual.EstadoNom = o("ESTADO")
                ordenActual.Estado = o("IdEstado")
                ordenActual.Servicio = o("IdServicio")
                ordenActual.ServicioNom = o("SERVICIO")
                ordenActual.Recinto = o("RECINTO")
                ordenActual.Modulo = o("MODULO")
                ordenActual.Piso = o("PISO")
                ordenActual.UnidadDestinoNom = o("UNIDAD_DESTINO")
                ordenActual.UnidadDestino = o("IdUnidad")
                ordenActual.DescripcionProblema = o("DESCRIPCION_PROBLEMA")
                ordenActual.Operador = o("NomCompleto").ToUpper()
                ordenActual.Solicitante = o("SOLICITANTE")
                ordenActual.CargoSolicitante = o("CARGO_SOLICITANTE")
                ordenActual.CargoSolicitanteId = o("IdCargoSol")
                WobjOrdenDetalle.NOrden = o("ORDEN")
                Dim ordenDetalle As DataSet = WobjOrdenDetalle.QryGetOrdenesDetalleById()
                If ordenDetalle.Tables(0).Rows.Count > 0 Then
                    For Each d In ordenDetalle.Tables(0).Rows
                        Dim detalle = New OrdenDetalle()
                        detalle.Id = d("ID")
                        detalle.NOrden = d("N_ORDEN")
                        detalle.Fecha = d("FECHA")
                        detalle.TecnicoUnidad = d("TECNICO_UNIDAD")
                        detalle.JefeUnidad = d("JEFE_UNIDAD")
                        detalle.SolucionProblema = d("SOLUCION_PROBLEMA")
                        'detalle.CostoTotal = d("COSTO_TOTAL")
                        'detalle.JefeServicio = d("JEFE_SERVICIO")
                        detalle.EstadoId = d("ESTADOID")
                        detalle.EstadoNom = d("ESTADO")
                        detalle.Usuario = d("USUARIO")
                        detalle.JefeServicio = d("SUPERVISIO_JEFE").ToString()
                        detalle.Comentario = d("COMENTARIO").ToString()
                        listaDetalleOt.Add(detalle)
                    Next

                End If

                WobjOrdenMateriales.NOrden = o("ORDEN")
                Dim ordenMateriales As DataSet = WobjOrdenMateriales.GetMaterialesByOt()
                If ordenMateriales.Tables(0).Rows.Count > 0 Then
                    For Each m In ordenMateriales.Tables(0).Rows
                        Dim Mat = New Materiales()
                        Mat.Id = m("id")
                        Mat.NOrden = m("OT")
                        Mat.MaterialNom = m("NomItem")
                        Mat.Cantidad = m("CantidadItem")
                        Mat.Precio = m("ValorItem")
                        listaMaterialesOt.Add(Mat)
                    Next
                End If
            Next

            'Dim oImagen As Image
            Const coordenadaX As Single = 15
            Const coordenadaY As Single = 700

            'Logo 

            Dim test As Drawing.Image = Drawing.Image.FromHbitmap(My.Resources.LogoMinisterio.GetHbitmap())
            Dim oImagen As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(test, Imaging.ImageFormat.Png)

            'oImagen = iTextSharp.text.Image.GetInstance(smileyPath)

            oImagen.SetAbsolutePosition(coordenadaX, coordenadaY)
            oImagen.ScaleToFit(100.0F, 100.0F)

            documento.Add(oImagen)

            documento.Add(New Paragraph(" "))
            documento.Add(New Paragraph(" "))
            documento.Add(New Paragraph(" "))

            Dim table As New PdfPTable(4)
            table.TotalWidth = 400.0F
            table.LockedWidth = True
            table.HorizontalAlignment = Element.ALIGN_RIGHT

            Dim medidaCel As Single() = {0.55F, 0.2F, 2.8F, 2.9F}

            table.HorizontalAlignment = Element.ALIGN_RIGHT
            table.SetWidths(medidaCel)

            Dim nested As New PdfPTable(2)
            Dim bottom As New PdfPCell(New Phrase("  ORDEN DE TRABAJO", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 18)))
            bottom.Colspan = 3
            bottom.Border = 0

            table.AddCell(bottom)
            Dim clTitulo = New PdfPCell(New Phrase("N° ORDEN: ", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clTitulo.BackgroundColor = BaseColor.LIGHT_GRAY
            'clTitulo.HorizontalAlignment = 1
            clTitulo.Padding = 3
            nested.AddCell(clTitulo)

            Dim clNumero = New PdfPCell(New Phrase(ordenActual.NOrden.ToString(), FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            nested.AddCell(clNumero)

            Dim clFechaText = New PdfPCell(New Phrase("FECHA :", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clFechaText.BackgroundColor = BaseColor.LIGHT_GRAY
            'clFechaText.HorizontalAlignment = 1
            clFechaText.Padding = 3
            nested.AddCell(clFechaText)

            Dim clFecha = New PdfPCell(New Phrase(Format(ordenActual.Fecha, "dd/MM/yyy"), FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            nested.AddCell(clFecha)

            Dim clHoraText = New PdfPCell(New Phrase("HORA :", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clHoraText.BackgroundColor = BaseColor.LIGHT_GRAY
            'clHoraText.HorizontalAlignment = 1
            clHoraText.Padding = 3
            nested.AddCell(clHoraText)

            Dim clHora = New PdfPCell(New Phrase(Format(ordenActual.Fecha, "HH:mm"), FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            nested.AddCell(clHora)
            Dim nesthousing As New PdfPCell(nested)
            nesthousing.Padding = 0.0F
            table.AddCell(nesthousing)

            documento.Add(table)
            documento.Add(New Paragraph(" "))
            documento.Add(New Paragraph(" "))

            '***********************DETALLE CREACION ORDEN******************************** 
            Dim tblDetalleCreacion As New PdfPTable(4)
            tblDetalleCreacion.TotalWidth = 560.0F
            tblDetalleCreacion.LockedWidth = True

            Dim clServicioText = New PdfPCell(New Phrase("SERVICIO:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clServicioText.BackgroundColor = BaseColor.LIGHT_GRAY
            clServicioText.Padding = 3
            tblDetalleCreacion.AddCell(clServicioText)

            Dim clServicio = New PdfPCell(New Phrase(ordenActual.ServicioNom, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            tblDetalleCreacion.AddCell(clServicio)

            Dim clSolicitanteText = New PdfPCell(New Phrase("SOLICITANTE:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clSolicitanteText.BackgroundColor = BaseColor.LIGHT_GRAY
            clSolicitanteText.Padding = 3
            tblDetalleCreacion.AddCell(clSolicitanteText)

            Dim clSolicitante = New PdfPCell(New Phrase(ordenActual.Solicitante, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            tblDetalleCreacion.AddCell(clSolicitante)

            Dim clRecintoText = New PdfPCell(New Phrase("RECINTO:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clRecintoText.BackgroundColor = BaseColor.LIGHT_GRAY
            clRecintoText.Padding = 3
            tblDetalleCreacion.AddCell(clRecintoText)

            Dim clRecinto = New PdfPCell(New Phrase(ordenActual.Recinto, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            tblDetalleCreacion.AddCell(clRecinto)

            Dim clUdesText = New PdfPCell(New Phrase("UNIDAD DESTINO:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clUdesText.BackgroundColor = BaseColor.LIGHT_GRAY
            clUdesText.Padding = 3
            tblDetalleCreacion.AddCell(clUdesText)

            Dim clUdes = New PdfPCell(New Phrase(ordenActual.UnidadDestinoNom, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            tblDetalleCreacion.AddCell(clUdes)

            Dim clModuloText = New PdfPCell(New Phrase("MODULO:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clModuloText.BackgroundColor = BaseColor.LIGHT_GRAY
            clModuloText.Padding = 3
            tblDetalleCreacion.AddCell(clModuloText)

            Dim clModulo = New PdfPCell(New Phrase(IIf(ordenActual.Modulo = "", "", ordenActual.Modulo), FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            tblDetalleCreacion.AddCell(clModulo)

            Dim clRecepText = New PdfPCell(New Phrase("RECEPCION:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clRecepText.BackgroundColor = BaseColor.LIGHT_GRAY
            clRecepText.Padding = 3
            tblDetalleCreacion.AddCell(clRecepText)

            Dim clRecep = New PdfPCell(New Phrase(ordenActual.Operador, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            tblDetalleCreacion.AddCell(clRecep)

            Dim clPisoText = New PdfPCell(New Phrase("PISO:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clPisoText.BackgroundColor = BaseColor.LIGHT_GRAY
            clPisoText.Padding = 3
            tblDetalleCreacion.AddCell(clPisoText)

            Dim clPiso = New PdfPCell(New Phrase(IIf(ordenActual.Piso = "", "", ordenActual.Piso), FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            tblDetalleCreacion.AddCell(clPiso)

            Dim clV1 = New PdfPCell(New Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            clV1.Border = 0
            tblDetalleCreacion.AddCell(clV1)

            Dim clV2 = New PdfPCell(New Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            clV2.Border = 0
            tblDetalleCreacion.AddCell(clV2)

            '***********************DESCRIPCION DEL PROBLEMA********************************       
            Dim tblDescProblema As New PdfPTable(1)
            tblDescProblema.TotalWidth = 560.0F
            tblDescProblema.LockedWidth = True

            Dim header1 = New PdfPCell(New Phrase("DESCRIPCION DEL PROBLEMA:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            header1.BackgroundColor = BaseColor.LIGHT_GRAY
            header1.Padding = 3
            header1.Colspan = 4
            header1.BorderWidthBottom = 0
            tblDescProblema.AddCell(header1)

            Dim desc = New PdfPCell(New Phrase(ordenActual.DescripcionProblema, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            desc.Colspan = 4
            desc.BorderWidthTop = 0
            tblDescProblema.AddCell(desc)

            '***********************FECHAS********************************
            Dim tblHoras As New PdfPTable(4)
            tblHoras.TotalWidth = 560.0F
            tblHoras.LockedWidth = True

            Dim clFechaIniText = New PdfPCell(New Phrase("FECHA INICIO:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clFechaIniText.BackgroundColor = BaseColor.LIGHT_GRAY
            clFechaIniText.Padding = 3
            tblHoras.AddCell(clFechaIniText)

            Dim ordenInicio As DataSet = WobjOrdenMateriales.GetInicioFinOt(ordenActual.NOrden, 1, "MIN")
            If ordenInicio.Tables(0).Rows.Count > 0 Then
                Dim clFechaIni = New PdfPCell(New Phrase(Format(ordenInicio.Tables(0).Rows(0).Item("FECHA"), "dd/MM/yyy"), FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clFechaIni.Padding = 2.0F
                tblHoras.AddCell(clFechaIni)
            Else
                Dim clFechaIni = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clFechaIni.Padding = 2.0F
                tblHoras.AddCell(clFechaIni)
            End If
            Dim clHoraIniText = New PdfPCell(New Phrase("HORA INICIO:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clHoraIniText.BackgroundColor = BaseColor.LIGHT_GRAY
            clHoraIniText.Padding = 3
            tblHoras.AddCell(clHoraIniText)

            If ordenInicio.Tables(0).Rows.Count > 0 Then
                Dim clHoraIni = New PdfPCell(New Phrase(Format(ordenInicio.Tables(0).Rows(0).Item("FECHA"), "HH:mm"), FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clHoraIni.Padding = 3
                tblHoras.AddCell(clHoraIni)
            Else
                Dim clHoraIni = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clHoraIni.Padding = 3
                tblHoras.AddCell(clHoraIni)
            End If

            Dim ordenFin As DataSet = WobjOrdenMateriales.GetInicioFinOt(ordenActual.NOrden, 3, "MAX")
            Dim clFechaTerText = New PdfPCell(New Phrase("FECHA TERMINO:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clFechaTerText.BackgroundColor = BaseColor.LIGHT_GRAY
            clFechaTerText.Padding = 3
            tblHoras.AddCell(clFechaTerText)

            If ordenFin.Tables(0).Rows.Count > 0 Then
                Dim clFechaTer = New PdfPCell(New Phrase(Format(ordenFin.Tables(0).Rows(0).Item("FECHA"), "dd/MM/yyy"), FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clFechaTer.Padding = 3
                tblHoras.AddCell(clFechaTer)
            Else
                Dim clFechaTer = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clFechaTer.Padding = 3
                tblHoras.AddCell(clFechaTer)
            End If

            Dim clHoraTerText = New PdfPCell(New Phrase("HORA TERMINO:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clHoraTerText.BackgroundColor = BaseColor.LIGHT_GRAY
            clHoraTerText.Padding = 3
            tblHoras.AddCell(clHoraTerText)

            If ordenFin.Tables(0).Rows.Count > 0 Then
                Dim clHoraTer = New PdfPCell(New Phrase(Format(ordenFin.Tables(0).Rows(0).Item("FECHA"), "HH:mm"), FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clHoraTer.Padding = 3
                tblHoras.AddCell(clHoraTer)
            Else
                Dim clHoraTer = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clHoraTer.Padding = 3
                tblHoras.AddCell(clHoraTer)
            End If
            '***********************TRABAJO REALIZADO********************************       
            Dim tblTrabrealizado As New PdfPTable(4)
            tblTrabrealizado.TotalWidth = 560.0F
            tblTrabrealizado.LockedWidth = True

            Dim clTituloTrabrealizado = New PdfPCell(New Phrase("TRABAJO REALIZADO:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clTituloTrabrealizado.BackgroundColor = BaseColor.LIGHT_GRAY
            clTituloTrabrealizado.Padding = 3
            clTituloTrabrealizado.Colspan = 4
            clTituloTrabrealizado.BorderWidthBottom = 0
            tblTrabrealizado.AddCell(clTituloTrabrealizado)


            If (listaDetalleOt.Count > 0) Then
                If ordenFin.Tables(0).Rows.Count > 0 Then
                    Dim clDetTrabrealizado = New PdfPCell(New Phrase(ordenFin.Tables(0).Rows(0).Item("SOLUCION_PROBLEMA"), FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    clDetTrabrealizado.Colspan = 4
                    clDetTrabrealizado.BorderWidthTop = 0
                    tblTrabrealizado.AddCell(clDetTrabrealizado)

                    Dim clV6Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    clV6Trabrealizado.Colspan = 3
                    clV6Trabrealizado.Border = 0
                    tblTrabrealizado.AddCell(clV6Trabrealizado)

                    'Dim clV4Trabrealizado = New PdfPCell(New Phrase("HORAS TOTAL: ", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
                    'tblTrabrealizado.AddCell(clV4Trabrealizado)
                Else
                    Dim clV1Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    clV1Trabrealizado.Colspan = 4
                    clV1Trabrealizado.BorderColorBottom = iTextSharp.text.BaseColor.WHITE
                    tblTrabrealizado.AddCell(clV1Trabrealizado)

                    Dim clV2Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    clV2Trabrealizado.Colspan = 4
                    tblTrabrealizado.AddCell(clV2Trabrealizado)

                    Dim clV3Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    clV3Trabrealizado.Colspan = 4
                    tblTrabrealizado.AddCell(clV3Trabrealizado)

                    Dim clV5Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    clV5Trabrealizado.Colspan = 4
                    tblTrabrealizado.AddCell(clV5Trabrealizado)

                    Dim clV6Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    clV6Trabrealizado.Colspan = 3
                    clV6Trabrealizado.Border = 0
                    tblTrabrealizado.AddCell(clV6Trabrealizado)

                    'Dim clV4Trabrealizado = New PdfPCell(New Phrase("HORAS TOTAL: ", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
                    'tblTrabrealizado.AddCell(clV4Trabrealizado)
                End If
            Else
                Dim clV1Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clV1Trabrealizado.Colspan = 4
                clV1Trabrealizado.BorderColorBottom = iTextSharp.text.BaseColor.WHITE
                tblTrabrealizado.AddCell(clV1Trabrealizado)

                Dim clV2Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clV2Trabrealizado.Colspan = 4
                tblTrabrealizado.AddCell(clV2Trabrealizado)

                Dim clV3Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clV3Trabrealizado.Colspan = 4
                tblTrabrealizado.AddCell(clV3Trabrealizado)

                Dim clV5Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clV5Trabrealizado.Colspan = 4
                tblTrabrealizado.AddCell(clV5Trabrealizado)

                Dim clV6Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clV6Trabrealizado.Colspan = 3
                clV6Trabrealizado.Border = 0
                tblTrabrealizado.AddCell(clV6Trabrealizado)

                'Dim clV4Trabrealizado = New PdfPCell(New Phrase("HORAS TOTAL: ", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
                'tblTrabrealizado.AddCell(clV4Trabrealizado)
            End If


            '***********************MATERIALES********************************       
            Dim tblMateriales As New PdfPTable(3)
            tblMateriales.TotalWidth = 560.0F
            tblMateriales.LockedWidth = True

            Dim clTituloMateriales = New PdfPCell(New Phrase("MATERIALES", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clTituloMateriales.Colspan = 3
            clTituloMateriales.Border = 0
            tblMateriales.AddCell(clTituloMateriales)

            Dim clCantida = New PdfPCell(New Phrase("CANTIDAD", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clCantida.BackgroundColor = BaseColor.LIGHT_GRAY
            clCantida.Padding = 3
            tblMateriales.AddCell(clCantida)

            Dim clArticulos = New PdfPCell(New Phrase("ARTICULOS Y DESCRIPCION", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clArticulos.BackgroundColor = BaseColor.LIGHT_GRAY
            clArticulos.Padding = 3
            tblMateriales.AddCell(clArticulos)

            Dim clPrecio = New PdfPCell(New Phrase("PRECIO UNITARIO", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clPrecio.BackgroundColor = BaseColor.LIGHT_GRAY
            clPrecio.Padding = 3
            tblMateriales.AddCell(clPrecio)
            Dim SumaTotal As [Decimal] = 0
            If (listaMaterialesOt.Count > 0) Then
                For Each materiales As Materiales In listaMaterialesOt
                    Dim cl1 = New PdfPCell(New Phrase(materiales.Cantidad, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    tblMateriales.AddCell(cl1)

                    Dim cl2 = New PdfPCell(New Phrase(materiales.MaterialNom, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    tblMateriales.AddCell(cl2)

                    Dim cl3 = New PdfPCell(New Phrase(materiales.Precio.ToString("c"), FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    tblMateriales.AddCell(cl3)
                    SumaTotal += Convert.ToDecimal(materiales.Precio) * materiales.Cantidad
                Next
                Dim clV6Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clV6Trabrealizado.Colspan = 2
                clV6Trabrealizado.Border = 0
                tblMateriales.AddCell(clV6Trabrealizado)

                Dim clV4Trabrealizado = New PdfPCell(New Phrase("TOTAL: " + SumaTotal.ToString("c"), FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
                tblMateriales.AddCell(clV4Trabrealizado)
            Else
                Dim cl1 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl1)
                Dim cl2 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl2)
                Dim cl3 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl3)

                Dim cl1A = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl1A)
                Dim cl2A = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl2A)
                Dim cl3A = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl3A)

                Dim cl1A1 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl1A1)
                Dim cl2A1 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl2A1)
                Dim cl3A1 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl3A1)

                Dim cl1A11 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl1A11)
                Dim cl2A11 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl2A11)
                Dim cl3A111 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl3A111)

                Dim cl1A1B = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl1A1B)
                Dim cl2A1B = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl2A1B)
                Dim cl3A1B = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblMateriales.AddCell(cl3A1B)

                Dim clV6Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clV6Trabrealizado.Colspan = 2
                clV6Trabrealizado.Border = 0
                tblMateriales.AddCell(clV6Trabrealizado)

                Dim clV4Trabrealizado = New PdfPCell(New Phrase("TOTAL: ", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
                tblMateriales.AddCell(clV4Trabrealizado)
            End If

            '***********************OBSERVACIONES********************************       
            Dim tblObser As New PdfPTable(4)
            tblObser.TotalWidth = 560.0F
            tblObser.LockedWidth = True

            Dim clTituloObser = New PdfPCell(New Phrase("OBSERVACIONES:", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            Comtario = listaDetalleOt(0).Comentario

            clTituloObser.BackgroundColor = BaseColor.LIGHT_GRAY
            clTituloObser.Padding = 3
            clTituloObser.Colspan = 4
            clTituloObser.BorderWidthBottom = 0
            tblObser.AddCell(clTituloObser)

            If (listaDetalleOt.Count > 0 And Comtario <> "") Then
                Dim clS1Trabrealizado = New PdfPCell(New Phrase(Comtario, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clS1Trabrealizado.Colspan = 4
                clS1Trabrealizado.BorderColorBottom = iTextSharp.text.BaseColor.WHITE
                tblObser.AddCell(clS1Trabrealizado)
            Else
                Dim clS1Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clS1Trabrealizado.Colspan = 4
                clS1Trabrealizado.BorderColorBottom = iTextSharp.text.BaseColor.WHITE
                tblObser.AddCell(clS1Trabrealizado)

                Dim clS2Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clS2Trabrealizado.Colspan = 4
                tblObser.AddCell(clS2Trabrealizado)

                Dim clS3Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clS3Trabrealizado.Colspan = 4
                tblObser.AddCell(clS3Trabrealizado)

                Dim clS5Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                clS5Trabrealizado.Colspan = 4
                tblObser.AddCell(clS5Trabrealizado)

            End If

            'Dim clS2Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            'clS2Trabrealizado.Colspan = 4
            'tblObser.AddCell(clS2Trabrealizado)

            'Dim clS3Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            'clS3Trabrealizado.Colspan = 4
            'tblObser.AddCell(clS3Trabrealizado)

            'Dim clS5Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            'clS5Trabrealizado.Colspan = 4
            'tblObser.AddCell(clS5Trabrealizado)

            Dim clS6Trabrealizado = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
            clS6Trabrealizado.Colspan = 3
            clS6Trabrealizado.Border = 0
            tblObser.AddCell(clS6Trabrealizado)


            '***********************RECEPCION CONFORME********************************       
            Dim tblRecepTitulo As New PdfPTable(1)
            tblRecepTitulo.TotalWidth = 560.0F
            tblRecepTitulo.LockedWidth = True

            Dim clRecepTitulo = New PdfPCell(New Phrase("RECEPCION CONFORME", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clRecepTitulo.Border = 0
            tblRecepTitulo.AddCell(clRecepTitulo)

            Dim tblRecepTituloTabla As New PdfPTable(3)
            tblRecepTituloTabla.TotalWidth = 560.0F
            tblRecepTituloTabla.LockedWidth = True
            tblRecepTituloTabla.DefaultCell.FixedHeight = 100.0F
            Dim ordenRecep As DataSet = WobjOrdenMateriales.GetInicioFinOt(ordenActual.NOrden, 3, "MIN")
            Dim ordenRecep1 As DataSet = WobjOrdenMateriales.GetInicioFinOt(ordenActual.NOrden, 2, "MIN")


            '********* Nuevo ********
            If ((ordenRecep.Tables(0).Rows.Count = 0 And ordenRecep1.Tables(0).Rows.Count = 0) And listaDetalleOt.Count > 0) Then
                Dim clRecepTituloTabla1 = New PdfPCell(New Phrase(listaDetalleOt(0).TecnicoUnidad, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblRecepTituloTabla.AddCell(clRecepTituloTabla1)
            End If
            '********* Fin Nuevo ********
            If (listaDetalleOt.Count > 0) Then
                If (ordenRecep.Tables(0).Rows.Count > 0) Then
                    Dim clRecepTituloTabla1 = New PdfPCell(New Phrase(listaDetalleOt(0).TecnicoUnidad, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    tblRecepTituloTabla.AddCell(clRecepTituloTabla1)
                Else
                    Dim clRecepTituloTabla1 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    tblRecepTituloTabla.AddCell(clRecepTituloTabla1)
                End If

                If (ordenRecep.Tables(0).Rows.Count > 0) Then
                    Dim clRecepTituloTabla2 = New PdfPCell(New Phrase(listaDetalleOt(0).JefeServicio, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    tblRecepTituloTabla.AddCell(clRecepTituloTabla2)
                Else
                    Dim clRecepTituloTabla2 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    tblRecepTituloTabla.AddCell(clRecepTituloTabla2)
                End If

                If (ordenRecep.Tables(0).Rows.Count > 0) Then
                    Dim clRecepTituloTabla3 = New PdfPCell(New Phrase(listaDetalleOt(0).JefeUnidad, FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    tblRecepTituloTabla.AddCell(clRecepTituloTabla3)
                Else
                    Dim clRecepTituloTabla3 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                    tblRecepTituloTabla.AddCell(clRecepTituloTabla3)
                End If
            Else
                Dim clRecepTituloTabla1 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblRecepTituloTabla.AddCell(clRecepTituloTabla1)

                Dim clRecepTituloTabla2 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblRecepTituloTabla.AddCell(clRecepTituloTabla2)

                Dim clRecepTituloTabla3 = New PdfPCell(New Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10)))
                tblRecepTituloTabla.AddCell(clRecepTituloTabla3)
            End If

            Dim tblJefes As New PdfPTable(3)
            tblJefes.TotalWidth = 560.0F
            tblJefes.LockedWidth = True

            Dim clJefes1 = New PdfPCell(New Phrase("TECNICOS", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clJefes1.Border = 0
            clJefes1.HorizontalAlignment = Element.ALIGN_CENTER
            tblJefes.AddCell(clJefes1)

            Dim clJefes2 = New PdfPCell(New Phrase("SUPERVISIOR O JEFE SERVICIO", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clJefes2.Border = 0
            clJefes2.HorizontalAlignment = Element.ALIGN_CENTER
            tblJefes.AddCell(clJefes2)

            Dim clJefes3 = New PdfPCell(New Phrase("JEFE UNIDAD", FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10)))
            clJefes3.Border = 0
            clJefes3.HorizontalAlignment = Element.ALIGN_CENTER
            tblJefes.AddCell(clJefes3)

            documento.Add(tblDetalleCreacion)
            documento.Add(New Paragraph(" "))
            documento.Add(tblDescProblema)
            documento.Add(New Paragraph(" "))
            documento.Add(tblHoras)
            documento.Add(New Paragraph(" "))
            documento.Add(tblTrabrealizado)
            documento.Add(New Paragraph(" "))
            documento.Add(tblMateriales)
            documento.Add(New Paragraph(" "))
            documento.Add(tblObser)
            documento.Add(New Paragraph(" "))
            documento.Add(tblRecepTitulo)
            documento.Add(tblRecepTituloTabla)
            documento.Add(tblJefes)

            documento.Close()
        End If
    End Function
End Class
