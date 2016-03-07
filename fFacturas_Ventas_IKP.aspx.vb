Imports System.Data
Partial Class fFacturas_Ventas_IKP
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT	Facturas.Cod_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Facturas.Nom_Cli END) END) AS  Nom_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Facturas.Rif = '') THEN Clientes.Rif ELSE Facturas.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Clientes.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Facturas.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Facturas.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Facturas.Telefonos = '') THEN Clientes.Telefonos ELSE Facturas.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("           Clientes.Fax, ")
            loConsulta.AppendLine("           Clientes.Generico, ")
            loConsulta.AppendLine("           Facturas.Nom_Cli        As  Nom_Gen, ")
            loConsulta.AppendLine("           Facturas.Rif            As  Rif_Gen, ")
            loConsulta.AppendLine("           Facturas.Nit            As  Nit_Gen, ")
            loConsulta.AppendLine("           Facturas.Dir_Fis        As  Dir_Gen, ")
            loConsulta.AppendLine("           Facturas.Telefonos      As  Tel_Gen, ")
            loConsulta.AppendLine("           Facturas.Documento, ")
            loConsulta.AppendLine("           Facturas.Fec_Ini, ")
            loConsulta.AppendLine("           Facturas.Fec_Fin, ")
            loConsulta.AppendLine("           Facturas.Mon_Bru, ")
            loConsulta.AppendLine("           Facturas.Mon_Imp1, ")
            loConsulta.AppendLine("           Facturas.Por_Imp1, ")
            loConsulta.AppendLine("           Facturas.Mon_Net, ")
            loConsulta.AppendLine("           Facturas.Por_Des1, ")
            loConsulta.AppendLine("           Facturas.Dis_Imp, ")
            loConsulta.AppendLine("           Facturas.Mon_Des1       As  Mon_Des, ")
            loConsulta.AppendLine("           Facturas.Por_Rec1, ")
            loConsulta.AppendLine("           Facturas.Mon_Rec1                       AS  Mon_Rec, ")
            loConsulta.AppendLine("           Facturas.Cod_For                        AS  Cod_For, ")
            loConsulta.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,25)    AS  Nom_For, ")
            loConsulta.AppendLine("           Facturas.Cod_Ven, ")
            loConsulta.AppendLine("           Facturas.Comentario, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Renglones_Facturas.Cod_Art, ")
            loConsulta.AppendLine("           CASE WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art ")
            loConsulta.AppendLine("			    ELSE Renglones_Facturas.Notas END AS Nom_Art,  ")
            loConsulta.AppendLine("           Renglones_Facturas.Renglon, ")
            loConsulta.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN Renglones_Facturas.Can_Art1")
            loConsulta.AppendLine("			    ELSE Renglones_Facturas.Can_Art2 END) AS Can_Art1, ")
            loConsulta.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN Renglones_Facturas.Cod_Uni")
            loConsulta.AppendLine("			    ELSE Renglones_Facturas.Cod_Uni2 END) AS Cod_Uni, ")
            loConsulta.AppendLine("           (CASE WHEN (Renglones_Facturas.Cod_Uni2='') THEN Renglones_Facturas.Precio1")
            loConsulta.AppendLine("			    ELSE Renglones_Facturas.Precio1*Renglones_Facturas.Can_Uni2 END) AS Precio1, ")
            loConsulta.AppendLine("           Renglones_Facturas.Mon_Net  As  Neto, ")
            loConsulta.AppendLine("           Renglones_Facturas.Por_Imp1 As  Por_Imp, ")
            loConsulta.AppendLine("           Renglones_Facturas.Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_Facturas.Mon_Imp1 As  Impuesto, ")
            loConsulta.AppendLine("           Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("           COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.direccion                     AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.telefonos                     AS Telefono_Empresa_Cliente  ")
            loConsulta.AppendLine(" FROM      Facturas ")
            loConsulta.AppendLine("   JOIN    Renglones_Facturas ON Facturas.Documento      =   Renglones_Facturas.Documento")
            loConsulta.AppendLine("   JOIN    Clientes ON Facturas.Cod_Cli      =   Clientes.Cod_Cli ")
            loConsulta.AppendLine("   JOIN    Formas_Pagos ON Facturas.Cod_For  =   Formas_Pagos.Cod_For")
            loConsulta.AppendLine("   JOIN    Vendedores ON Facturas.Cod_Ven    =   Vendedores.Cod_Ven")
            loConsulta.AppendLine("   JOIN    Articulos ON Articulos.Cod_Art   =   Renglones_Facturas.Cod_Art")
            loConsulta.AppendLine("  JOIN     Sucursales ON Sucursales.cod_suc = Facturas.cod_suc")
            loConsulta.AppendLine("  LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("        AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("        AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
      
           
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFacturas_Ventas_IKP", laDatosReporte)
			lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".",",")
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFacturas_Ventas_IKP.ReportSource = loObjetoReporte

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
' RJG: 01/08/13: Codigo inicial
'-------------------------------------------------------------------------------------------'
