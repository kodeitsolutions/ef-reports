'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCotizaciones_Clientes_IKP"
'-------------------------------------------------------------------------------------------'
Partial Class fCotizaciones_Clientes_IKP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT	   Cotizaciones.Cod_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN Clientes.Nom_Cli ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cotizaciones.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Cotizaciones.Nom_Cli END) END) AS  Nom_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN Clientes.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cotizaciones.Rif = '') THEN Clientes.Rif ELSE Cotizaciones.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Clientes.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Cotizaciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Cotizaciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND CAST (Cotizaciones.Nom_Cli AS VARCHAR) = '') THEN Clientes.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cotizaciones.Telefonos = '') THEN Clientes.Telefonos ELSE Cotizaciones.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("           Clientes.Fax, ")
            loConsulta.AppendLine("           Cotizaciones.Nom_Cli                    As  Nom_Gen, ")
            loConsulta.AppendLine("           Cotizaciones.Rif                        As  Rif_Gen, ")
            loConsulta.AppendLine("           Cotizaciones.Nit                        As  Nit_Gen, ")
            loConsulta.AppendLine("           Cotizaciones.Dir_Fis                    As  Dir_Gen, ")
            loConsulta.AppendLine("           Cotizaciones.Telefonos                  As  Tel_Gen, ")
            loConsulta.AppendLine("           Cotizaciones.Documento, ")
            loConsulta.AppendLine("           Cotizaciones.Fec_Ini, ")
            loConsulta.AppendLine("           Cotizaciones.Fec_Fin, ")
            loConsulta.AppendLine("           Cotizaciones.Mon_Bru, ")
            loConsulta.AppendLine("           Cotizaciones.Por_Des1, ")
            loConsulta.AppendLine("           Cotizaciones.Por_Rec1, ")
            loConsulta.AppendLine("           Cotizaciones.Mon_Des1, ")
            loConsulta.AppendLine("           Cotizaciones.Mon_Rec1, ")
            loConsulta.AppendLine("           Cotizaciones.Mon_Imp1, ")
            loConsulta.AppendLine("           Cotizaciones.Mon_Net, ")
            loConsulta.AppendLine("           Cotizaciones.Cod_For, ")
            loConsulta.AppendLine("           Cotizaciones.Dis_Imp, ")
            loConsulta.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loConsulta.AppendLine("           Cotizaciones.Cod_Ven, ")
            loConsulta.AppendLine("           Cotizaciones.Comentario, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Cod_Art, ")


            loConsulta.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            loConsulta.AppendLine("			    ELSE Renglones_Cotizaciones.Notas END AS Nom_Art,  ")

            loConsulta.AppendLine("           Renglones_Cotizaciones.Renglon, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Comentario AS Comentario_renglon, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Can_Art1, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Cod_Uni, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Precio1, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Por_Des, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Mon_Net           As  Neto, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Por_Imp1          As  Por_Imp, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_Cotizaciones.Mon_Imp1          AS  Impuesto, ")
            loConsulta.AppendLine("           Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("           COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.direccion                     AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.telefonos                     AS Telefono_Empresa_Cliente   ")
            loConsulta.AppendLine("FROM       Cotizaciones ")
            loConsulta.AppendLine("    JOIN   Renglones_Cotizaciones ON Cotizaciones.Documento = Renglones_Cotizaciones.Documento")
            loConsulta.AppendLine("    JOIN   Clientes ON Cotizaciones.Cod_Cli = Clientes.Cod_Cli")
            loConsulta.AppendLine("    JOIN   Formas_Pagos ON Cotizaciones.Cod_For = Formas_Pagos.Cod_For")
            loConsulta.AppendLine("    JOIN   Vendedores ON Cotizaciones.Cod_Ven = Vendedores.Cod_Ven")
            loConsulta.AppendLine("    JOIN   Articulos ON Articulos.Cod_Art = Renglones_Cotizaciones.Cod_Art")
            loConsulta.AppendLine("    JOIN   Sucursales ON Sucursales.cod_suc = Cotizaciones.cod_suc")
            loConsulta.AppendLine("    LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("            AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("            AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loConsulta.ToString())
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
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".", ",")


			'--------------------------------------------------'
			' Carga la imagen del logo en cusReportes          '
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCotizaciones_Clientes_IKP", laDatosReporte)
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text38"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCotizaciones_Clientes_IKP.ReportSource = loObjetoReporte

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
' RJG: 02/08/13: Codigo inicial
'-------------------------------------------------------------------------------------------'
