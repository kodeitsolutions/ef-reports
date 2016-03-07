Imports System.Data
Partial Class fOrdenes_Compras_IKP
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT     Ordenes_Compras.Cod_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Ordenes_Compras.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Compras.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Proveedores.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Ordenes_Compras.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Compras.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Ordenes_Compras.Telefonos = '') THEN Proveedores.Telefonos ELSE Ordenes_Compras.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("           Proveedores.Fax, ")
            loConsulta.AppendLine("           Ordenes_Compras.Nom_Pro         AS Nom_Gen, ")
            loConsulta.AppendLine("           Ordenes_Compras.Rif             AS Rif_Gen, ")
            loConsulta.AppendLine("           Ordenes_Compras.Nit             AS Nit_Gen, ")
            loConsulta.AppendLine("           Ordenes_Compras.Dir_Fis         AS Dir_Gen, ")
            loConsulta.AppendLine("           Ordenes_Compras.Telefonos       AS Tel_Gen, ")
            loConsulta.AppendLine("           Ordenes_Compras.Documento, ")
            
            loConsulta.AppendLine("           Ordenes_Compras.Por_Des1 AS Por_Des1_Enc, ")
            loConsulta.AppendLine("           Ordenes_Compras.Mon_Des1 AS Mon_Des1_Enc, ")
            loConsulta.AppendLine("           Ordenes_Compras.Por_Rec1 AS Por_Rec1_Enc, ")
            loConsulta.AppendLine("           Ordenes_Compras.Mon_Rec1 AS Mon_Rec1_Enc, ")
            
            loConsulta.AppendLine("           Renglones_OCompras.Cod_Uni, ")
            loConsulta.AppendLine("           Ordenes_Compras.Fec_Ini, ")
            loConsulta.AppendLine("           Ordenes_Compras.Fec_Fin, ")
            loConsulta.AppendLine("           Ordenes_Compras.Mon_Bru, ")
            loConsulta.AppendLine("           Ordenes_Compras.Por_Imp1, ")
            loConsulta.AppendLine("           Ordenes_Compras.Dis_Imp, ")
            loConsulta.AppendLine("           Ordenes_Compras.Mon_Imp1, ")
            loConsulta.AppendLine("           Ordenes_Compras.Mon_Net, ")
            loConsulta.AppendLine("           Ordenes_Compras.Cod_For, ")
            loConsulta.AppendLine("           Ordenes_Compras.Comentario, ")
            loConsulta.AppendLine("           Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("           Ordenes_Compras.Cod_Ven, ")
            loConsulta.AppendLine("           Renglones_OCompras.Cod_Art, ")
            loConsulta.AppendLine("		        CASE")
			loConsulta.AppendLine("		    	    WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
			loConsulta.AppendLine("		    	    ELSE Renglones_OCompras.Notas")
			loConsulta.AppendLine("		        END									AS Nom_Art,  ")            
            loConsulta.AppendLine("           Renglones_OCompras.Renglon, ")
            loConsulta.AppendLine("           Renglones_OCompras.Can_Art1, ")
            loConsulta.AppendLine("           Renglones_OCompras.Por_Des               AS Por_Des1, ")
            loConsulta.AppendLine("           Renglones_OCompras.Precio1               AS Precio1, ")
            loConsulta.AppendLine("           Renglones_OCompras.Precio1               AS Precio1, ")
            loConsulta.AppendLine("           Renglones_OCompras.Comentario            AS Comentario_Renglon, ")
            loConsulta.AppendLine("           Renglones_OCompras.Mon_Net               AS Neto, ")
            loConsulta.AppendLine("           Renglones_OCompras.Cod_Imp               AS Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_OCompras.Por_Imp1              AS Por_Imp, ")
            loConsulta.AppendLine("           Renglones_OCompras.Mon_Imp1              AS Impuesto, ")
            loConsulta.AppendLine("           Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("           COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.direccion                     AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.telefonos                     AS Telefono_Empresa_Cliente  ")
            loConsulta.AppendLine("FROM       Ordenes_Compras ")
            loConsulta.AppendLine("  JOIN     Renglones_OCompras ON Ordenes_Compras.Documento = Renglones_OCompras.Documento")
            loConsulta.AppendLine("  JOIN     Proveedores ON Ordenes_Compras.Cod_Pro = Proveedores.Cod_Pro ")
            loConsulta.AppendLine("  JOIN     Formas_Pagos ON Ordenes_Compras.Cod_For = Formas_Pagos.Cod_For ")
            loConsulta.AppendLine("  JOIN     Articulos ON Articulos.Cod_Art = Renglones_OCompras.Cod_Art")
            loConsulta.AppendLine("  JOIN     Sucursales ON Sucursales.cod_suc = ordenes_compras.cod_suc")
            loConsulta.AppendLine("  LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("        AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("        AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine("WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)

			
			
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")
            
			Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpueto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpueto = "("

            'Recorre cada renglon de la tabla
            For lnNumeroFila As Integer = 0 To laDatosReporte.Tables(0).Rows.Count - 1
                lcXml = laDatosReporte.Tables(0).Rows(lnNumeroFila).Item("dis_imp")

                If String.IsNullOrEmpty(lcXml.Trim()) Then
                    Continue For
                End If

                loImpuestos.LoadXml(lcXml)

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
						If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
								If CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText)<> 0 Then
									lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
								End If
						End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(,", "(")
            
            
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
			
			 '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOrdenes_Compras_IKP", laDatosReporte)
			lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".",",")
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text25"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfOrdenes_Compras_IKP.ReportSource = loObjetoReporte

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
' RJG: 01/08/13: Programacion inicial
'-------------------------------------------------------------------------------------------'
