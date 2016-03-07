Imports System.Data
Partial Class fRecepciones_Proveedores_IKP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT	Recepciones.Cod_Pro                    As  Cod_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Recepciones.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Recepciones.Nom_Pro END) END) AS  Nom_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Recepciones.Rif = '') THEN Proveedores.Rif ELSE Recepciones.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Proveedores.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Recepciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Recepciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Recepciones.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Recepciones.Telefonos = '') THEN Proveedores.Telefonos ELSE Recepciones.Telefonos END) END) AS  Telefonos, ")
            
            loConsulta.AppendLine("           Proveedores.Fax, ")
            loConsulta.AppendLine("           Recepciones.Nom_Pro                    As  Nom_Gen, ")
            loConsulta.AppendLine("           Recepciones.Rif                        As  Rif_Gen, ")
            loConsulta.AppendLine("           Recepciones.Nit                        As  Nit_Gen, ")
            loConsulta.AppendLine("           Recepciones.Dir_Fis                    As  Dir_Gen, ")
            loConsulta.AppendLine("           Recepciones.Telefonos                  As  Tel_Gen, ")
            loConsulta.AppendLine("           Recepciones.Documento, ")
            loConsulta.AppendLine("           Recepciones.Fec_Ini, ")
            loConsulta.AppendLine("           Recepciones.Fec_Fin, ")
            loConsulta.AppendLine("           Recepciones.Mon_Bru, ")
            loConsulta.AppendLine("           Recepciones.Mon_Imp1, ")
            loConsulta.AppendLine("           Recepciones.Dis_Imp, ")
            loConsulta.AppendLine("           Recepciones.Mon_Net, ")
            loConsulta.AppendLine("           Recepciones.Mon_Des1, ")
            loConsulta.AppendLine("           Recepciones.Mon_Rec1, ")
            loConsulta.AppendLine("           Recepciones.Por_Des1, ")
            loConsulta.AppendLine("           Recepciones.Por_Rec1, ")
            loConsulta.AppendLine("           Recepciones.Cod_For, ")
            loConsulta.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,30)      AS  Nom_For, ")
            loConsulta.AppendLine("           Recepciones.Cod_Ven, ")
            loConsulta.AppendLine("           Recepciones.Comentario, ")
            loConsulta.AppendLine("           SUBSTRING(Vendedores.Nom_Ven,1,25)        AS  Nom_Ven, ")
            loConsulta.AppendLine("           Renglones_Recepciones.Cod_Art, ")
            loConsulta.AppendLine("		CASE")
			loConsulta.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
			loConsulta.AppendLine("			ELSE Renglones_Recepciones.Notas")
			loConsulta.AppendLine("		END												AS Nom_Art,  ")            
            loConsulta.AppendLine("           Renglones_Recepciones.Renglon, ")
            loConsulta.AppendLine("           Renglones_Recepciones.Can_Art1, ")
            loConsulta.AppendLine("           Renglones_Recepciones.Cod_Uni, ")
            loConsulta.AppendLine("           Renglones_Recepciones.Precio1, ")
            loConsulta.AppendLine("           Renglones_Recepciones.Mon_Net              AS Neto, ")
            loConsulta.AppendLine("           Renglones_Recepciones.Por_Imp1             AS Por_Imp, ")
            loConsulta.AppendLine("           Renglones_Recepciones.Por_Des              AS Por_Des, ")
            loConsulta.AppendLine("           Renglones_Recepciones.Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_Recepciones.Mon_Imp1             AS Impuesto, ")
            loConsulta.AppendLine("           Sucursales.nom_suc                         AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("           COALESCE(campos_propiedades.val_car, '')   AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.direccion                       AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.telefonos                       AS Telefono_Empresa_Cliente   ")  
            loConsulta.AppendLine(" FROM      Recepciones ")
            loConsulta.AppendLine("    JOIN   Renglones_Recepciones ON Recepciones.Documento  =   Renglones_Recepciones.Documento")
            loConsulta.AppendLine("    JOIN   Proveedores ON Recepciones.Cod_Pro     = Proveedores.Cod_Pro")
            loConsulta.AppendLine("    JOIN   Formas_Pagos ON Recepciones.Cod_For    = Formas_Pagos.Cod_For")
            loConsulta.AppendLine("    JOIN   Vendedores ON Recepciones.Cod_Ven      = Vendedores.Cod_Ven")
            loConsulta.AppendLine("    JOIN   Articulos ON Articulos.Cod_Art         = Renglones_Recepciones.Cod_Art")
            loConsulta.AppendLine("    JOIN   Sucursales ON Sucursales.cod_suc       = Recepciones.cod_suc")
            loConsulta.AppendLine("    LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("            AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("            AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine(" WHERE    " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")

            
            Dim loServicios As New cusDatos.goDatos
            'Me.loConsulta(loComandoSeleccionar.ToString())
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
            
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


			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fRecepciones_Proveedores_IKP", laDatosReporte)
			lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".",",")          
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfRecepciones_Proveedores_IKP.ReportSource = loObjetoReporte

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
' JJD: 27/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 10/09/09: Se ajusto el nombre del articulo para los casos de aquellos articulos genericos
'-------------------------------------------------------------------------------------------'
' CMS: 16/09/09: Se Agrego la distribucion de impuesto
'-------------------------------------------------------------------------------------------'
' MAT: 01/03/11: Ajuste de la vista de diseño												'
'-------------------------------------------------------------------------------------------'