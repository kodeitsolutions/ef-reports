'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPedidos_Clientes_IKP"
'-------------------------------------------------------------------------------------------'
Partial Class fPedidos_Clientes_IKP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT	    Pedidos.Cod_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Nom_Cli ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Pedidos.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Pedidos.Nom_Cli END) END) AS  Nom_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Pedidos.Rif = '') THEN Clientes.Rif ELSE Pedidos.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Clientes.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Pedidos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Pedidos.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0) THEN Clientes.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Pedidos.Telefonos = '') THEN Clientes.Telefonos ELSE Pedidos.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("           Clientes.Fax, ")
            loConsulta.AppendLine("           Pedidos.Nom_Cli                    As  Nom_Gen, ")
            loConsulta.AppendLine("           Pedidos.Rif                        As  Rif_Gen, ")
            loConsulta.AppendLine("           Pedidos.Nit                        As  Nit_Gen, ")
            loConsulta.AppendLine("           Pedidos.Dir_Fis                    As  Dir_Gen, ")
            loConsulta.AppendLine("           Pedidos.Telefonos                  As  Tel_Gen, ")
            loConsulta.AppendLine("           Pedidos.Documento, ")
            loConsulta.AppendLine("           Pedidos.Fec_Ini, ")
            loConsulta.AppendLine("           Pedidos.Fec_Fin, ")
            loConsulta.AppendLine("           Pedidos.Mon_Bru, ")
            loConsulta.AppendLine("           Pedidos.Por_Des1, ")
            loConsulta.AppendLine("           Pedidos.Por_Rec1, ")
            loConsulta.AppendLine("           Pedidos.Mon_Des1, ")
            loConsulta.AppendLine("           Pedidos.Mon_Rec1, ")
            loConsulta.AppendLine("           Pedidos.Mon_Imp1, ")
            loConsulta.AppendLine("           Pedidos.Dis_Imp, ")
            loConsulta.AppendLine("           Pedidos.Por_Imp1, ")
            loConsulta.AppendLine("           Pedidos.Mon_Net, ")
            loConsulta.AppendLine("           Pedidos.Cod_For, ")
            loConsulta.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,25)    AS  Nom_For, ")
            loConsulta.AppendLine("           Pedidos.Cod_Ven, ")
            loConsulta.AppendLine("           Pedidos.Comentario, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Cod_Art, ")
            'loComandoSeleccionar.AppendLine("           Articulos.Nom_Art, ")
            loConsulta.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            loConsulta.AppendLine("			    ELSE Renglones_Pedidos.Notas END AS Nom_Art,  ")
            loConsulta.AppendLine("           Renglones_Pedidos.Renglon, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Can_Art1, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Cod_Uni, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Precio1, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Mon_Net          As  Neto, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Por_Imp1         As  Por_Imp, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_Pedidos.Mon_Imp1         As  Impuesto, ")
            loConsulta.AppendLine("           Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("           COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.direccion                     AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.telefonos                     AS Telefono_Empresa_Cliente   ")
            loConsulta.AppendLine("FROM       Pedidos ")
            loConsulta.AppendLine("    JOIN   Renglones_Pedidos ON Pedidos.Documento = Renglones_Pedidos.Documento")
            loConsulta.AppendLine("    JOIN   Clientes ON Pedidos.Cod_Cli = Clientes.Cod_Cli")
            loConsulta.AppendLine("    JOIN   Formas_Pagos ON Pedidos.Cod_For = Formas_Pagos.Cod_For")
            loConsulta.AppendLine("    JOIN   Vendedores ON Pedidos.Cod_Ven = Vendedores.Cod_Ven")
            loConsulta.AppendLine("    JOIN   Articulos ON Articulos.Cod_Art = Renglones_Pedidos.Cod_Art")
            loConsulta.AppendLine("    JOIN   Sucursales ON Sucursales.cod_suc = Pedidos.cod_suc")
            loConsulta.AppendLine("  LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("        AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("        AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine("WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")
            
			Dim lcXml As String = "<impuesto></impuesto>"
            Dim lcPorcentajesImpuesto As String
            Dim loImpuestos As New System.Xml.XmlDocument()

            lcPorcentajesImpuesto = "("

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
									lcPorcentajesImpuesto = lcPorcentajesImpuesto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
								End If
						End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpuesto = lcPorcentajesImpuesto & ")"
            lcPorcentajesImpuesto = lcPorcentajesImpuesto.Replace("(,", "(")
            
            
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

			
			
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPedidos_Clientes_IKP", laDatosReporte)
            lcPorcentajesImpuesto = lcPorcentajesImpuesto.Replace(".",",")
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpuesto.ToString

            
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPedidos_Clientes_IKP.ReportSource = loObjetoReporte

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
