Imports System.Data
Partial Class fPresupuestos_Proveedores_IKP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT	Presupuestos.Cod_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Presupuestos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Presupuestos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Presupuestos.Nom_Pro END) END) AS  Nom_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Presupuestos.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Presupuestos.Rif = '') THEN Proveedores.Rif ELSE Presupuestos.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Proveedores.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Presupuestos.Nom_Pro = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Presupuestos.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Presupuestos.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Presupuestos.Nom_Pro = '') THEN Proveedores.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Presupuestos.Telefonos = '') THEN Proveedores.Telefonos ELSE Presupuestos.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("           Proveedores.Fax, ")
            loConsulta.AppendLine("           Presupuestos.Nom_Pro                    As  Nom_Gen, ")
            loConsulta.AppendLine("           Presupuestos.Rif                        As  Rif_Gen, ")
            loConsulta.AppendLine("           Presupuestos.Nit                        As  Nit_Gen, ")
            loConsulta.AppendLine("           Presupuestos.Dir_Fis                    As  Dir_Gen, ")
            loConsulta.AppendLine("           Presupuestos.Telefonos                  As  Tel_Gen, ")
            loConsulta.AppendLine("           Presupuestos.Documento, ")
            loConsulta.AppendLine("           Presupuestos.Fec_Ini, ")
            loConsulta.AppendLine("           Presupuestos.Fec_Fin, ")
            loConsulta.AppendLine("           Presupuestos.Mon_Bru, ")
            loConsulta.AppendLine("           Presupuestos.Mon_Imp1, ")
            loConsulta.AppendLine("           Presupuestos.Por_Des1, ")
            loConsulta.AppendLine("           Presupuestos.Mon_Des1, ")
            loConsulta.AppendLine("           Presupuestos.Por_Rec1, ")
            loConsulta.AppendLine("           Presupuestos.Mon_Rec1, ")
            loConsulta.AppendLine("           Presupuestos.Dis_Imp, ")
            loConsulta.AppendLine("           Presupuestos.Mon_Net, ")
            loConsulta.AppendLine("           Presupuestos.Cod_For, ")
            loConsulta.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,20)    AS  Nom_For, ")
            loConsulta.AppendLine("           Presupuestos.Cod_Ven, ")
            loConsulta.AppendLine("           Presupuestos.Comentario, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Renglones_Presupuestos.Cod_Art, ")
            loConsulta.AppendLine("		CASE")
			loConsulta.AppendLine("			WHEN Articulos.Generico = 0 THEN Articulos.Nom_Art")
			loConsulta.AppendLine("			ELSE Renglones_Presupuestos.Notas")
			loConsulta.AppendLine("		END														AS Nom_Art,  ")            
            loConsulta.AppendLine("           Renglones_Presupuestos.Renglon, ")
            loConsulta.AppendLine("           Renglones_Presupuestos.Can_Art1, ")
            loConsulta.AppendLine("           Renglones_Presupuestos.Por_Des, ")
            loConsulta.AppendLine("           Renglones_Presupuestos.Cod_Uni, ")
            loConsulta.AppendLine("           Renglones_Presupuestos.Precio1, ")
            loConsulta.AppendLine("           Renglones_Presupuestos.Mon_Net          As  Neto, ")
            loConsulta.AppendLine("           Renglones_Presupuestos.Por_Imp1         As  Por_Imp, ")
            loConsulta.AppendLine("           Renglones_Presupuestos.Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_Presupuestos.Mon_Imp1         As  Impuesto, ")
            loConsulta.AppendLine("           Sucursales.nom_suc                         AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("           COALESCE(campos_propiedades.val_car, '')   AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.direccion                       AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.telefonos                       AS Telefono_Empresa_Cliente,  ") 
            loConsulta.AppendLine("           Sucursales.Fax                             AS Fax_Empresa_Cliente   ") 
			loConsulta.AppendLine(" FROM      Presupuestos ")
            loConsulta.AppendLine("       JOIN Renglones_Presupuestos on Presupuestos.Documento  =   Renglones_Presupuestos.Documento")
            loConsulta.AppendLine("       JOIN Proveedores ON Presupuestos.Cod_Pro    =   Proveedores.Cod_Pro ")
            loConsulta.AppendLine("       JOIN Formas_Pagos ON Presupuestos.Cod_For    =   Formas_Pagos.Cod_For ")
            loConsulta.AppendLine("       LEFT JOIN Vendedores ON Presupuestos.Cod_Ven    =   Vendedores.Cod_Ven ")
            loConsulta.AppendLine("       JOIN Articulos ON Articulos.Cod_Art       =   Renglones_Presupuestos.Cod_Art")
            loConsulta.AppendLine("    JOIN   Sucursales ON Sucursales.cod_suc  = Presupuestos.cod_suc")
            loConsulta.AppendLine("    LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("            AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("            AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine(" WHERE     " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")



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

                'En cada renglón lee el contenido de la distribución de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
                    If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
                    'Verifica si el impuesto es igual a Cero
 						if CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) <> 0 Then
							lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
						End If
                    End If
                Next loImpuesto
            Next lnNumeroFila

            lcPorcentajesImpueto = lcPorcentajesImpueto & ")"
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace("(,","(")
            
            if lcPorcentajesImpueto = "()" Then
					lcPorcentajesImpueto = " "
			End If

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

            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPresupuestos_Proveedores_IKP", laDatosReporte)
            
            lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".",",")
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPresupuestos_Proveedores_IKP.ReportSource = loObjetoReporte

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
'				Se modifico la consulta para hacer left join con vendedores
'-------------------------------------------------------------------------------------------'
' CMS: 16/09/09: Se Agrego la distribucion de impuesto
'-------------------------------------------------------------------------------------------'
' CMS: 11/06/10: Se coloco la validación de registros 0 y el metodo de carga de imagen 
'					Se agrego el descuento del renglon, proveedor generico, y el descuento y
'					recargo del documento
'-------------------------------------------------------------------------------------------' 
' MAT: 10/11/10: Mantenimiento del Reporte
'-------------------------------------------------------------------------------------------'