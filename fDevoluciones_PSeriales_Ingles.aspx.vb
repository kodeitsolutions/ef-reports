Imports System.Data

Partial Class fDevoluciones_Proveedores_Ingles

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT	   Devoluciones_Proveedores.Cod_Pro, ")
            loConsulta.AppendLine("            Proveedores.Nom_Pro, ")
            loConsulta.AppendLine("            Proveedores.Rif, ")
            loConsulta.AppendLine("            Proveedores.Nit, ")
            loConsulta.AppendLine("            Proveedores.Dir_Fis, ")
            loConsulta.AppendLine("            Proveedores.Telefonos, ")
            loConsulta.AppendLine("            Proveedores.Fax, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Nom_Pro        As  Nom_Gen, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Rif            As  Rif_Gen, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Nit            As  Nit_Gen, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Dir_Fis        As  Dir_Gen, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Telefonos      As  Tel_Gen, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Documento, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Fec_Ini, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Fec_Fin, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Mon_Bru, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Mon_Imp1, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Dis_Imp, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Mon_Net, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Por_Rec1, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Por_Des1, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Mon_Rec1, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Mon_Des1, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Cod_For, ")
            loConsulta.AppendLine("            Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Cod_Ven, ")
            loConsulta.AppendLine("            Devoluciones_Proveedores.Comentario, ")
            loConsulta.AppendLine("            Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("            Renglones_DProveedores.Cod_Art, ")
            loConsulta.AppendLine("            Articulos.Nom_Art, ")
            loConsulta.AppendLine("            Renglones_DProveedores.Renglon, ")
            loConsulta.AppendLine("            Renglones_DProveedores.Can_Art1,")
            loConsulta.AppendLine("            Renglones_DProveedores.Cod_Uni, ")
            loConsulta.AppendLine("            Renglones_DProveedores.Precio1, ")
            loConsulta.AppendLine("            Renglones_DProveedores.Mon_Net  As  Neto, ")
            loConsulta.AppendLine("            Renglones_DProveedores.Por_Imp1 As  Por_Imp, ")
            loConsulta.AppendLine("            Renglones_DProveedores.Cod_Imp, ")
            loConsulta.AppendLine("            Renglones_DProveedores.Mon_Imp1 As  Impuesto,")
            loConsulta.AppendLine("            Seriales.serial,")
            loConsulta.AppendLine("            REPLACE(COALESCE(seriales.tip_sal, ''), '_', ' ') AS Tipo_Serial")
            loConsulta.AppendLine("FROM        Devoluciones_Proveedores ")
            loConsulta.AppendLine("   JOIN     Renglones_DProveedores  ON Renglones_DProveedores.Documento = Devoluciones_Proveedores.Documento")
            loConsulta.AppendLine("   JOIN     Proveedores             ON Proveedores.Cod_Pro = Devoluciones_Proveedores.Cod_Pro")
            loConsulta.AppendLine("   JOIN     Formas_Pagos            ON Formas_Pagos.Cod_For = Devoluciones_Proveedores.Cod_For")
            loConsulta.AppendLine("   JOIN     Vendedores              ON Vendedores.Cod_Ven = Devoluciones_Proveedores.Cod_Ven")
            loConsulta.AppendLine("   JOIN     Articulos               ON Articulos.Cod_Art = Renglones_DProveedores.Cod_Art")
            loConsulta.AppendLine("   JOIN     Seriales                ON seriales.ren_sal = Renglones_DProveedores.renglon")
            loConsulta.AppendLine("                                   AND seriales.doc_sal = Renglones_DProveedores.Documento")
            loConsulta.AppendLine("                                   AND seriales.tip_sal ='Devoluciones_Proveedores' ")
            loConsulta.AppendLine("WHERE       " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("ORDER BY    Renglones_DProveedores.Renglon, seriales.renglon;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            
            'Me.mEscribirConsulta(loConsulta.ToString())


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            'Dim lcXml As String = "<impuesto></impuesto>"
            'Dim llEsDetalle AS Boolean = False
            'Dim lcPorcentajesImpuestos As String
            'Dim loImpuestos As New System.Xml.XmlDocument()

            'lcPorcentajesImpuestos = "("

            ''Recorre cada renglon de la tabla
            'For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
            '    lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("dis_imp")
            '    llEsDetalle = CBool(CInt(laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("contador"))=1)

            '    If Not llEsDetalle OrElse String.IsNullOrEmpty(lcXml.Trim()) Then
            '        Continue For
            '    End If

            '    loImpuestos.LoadXml(lcXml)

            '    'En cada renglón lee el contenido de la distribución de impuestos
            '    For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
            '        If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
            '            lcPorcentajesImpuestos = lcPorcentajesImpuestos & ", " & goServicios.mObtenerFormatoCadena(CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText)) & "%"
            '        End If
            '    Next loImpuesto
            'Next lnNumeroFila

            'lcPorcentajesImpuestos = lcPorcentajesImpuestos & ")"
            'lcPorcentajesImpuestos = lcPorcentajesImpuestos.Replace("(,", "(")


            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fDevoluciones_Proveedores_Ingles", laDatosReporte)

            'CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpuestos.ToString

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfDevoluciones_Proveedores_Ingles.ReportSource = loObjetoReporte


        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload

        Try

            loObjetoReporte.Close()

        Catch loExcepcion As Exception

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' Douglas Cortez: 10/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Ajuste de la vista de Diseño
'-------------------------------------------------------------------------------------------'
' MAT: 15/09/11: Eliminación del Pie de Página de eFactory según Requerimientos
'-------------------------------------------------------------------------------------------'
