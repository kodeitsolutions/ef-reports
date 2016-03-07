Imports System.Data
Partial Class fRequisiciones_Internas_IKP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT	Requisiciones.Cod_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cast(Requisiciones.Nom_Pro As Varchar)= '') THEN Proveedores.Nom_Pro ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Cast(Requisiciones.Nom_Pro As Varchar) = '') THEN Proveedores.Nom_Pro ELSE Requisiciones.Nom_Pro END) END) AS  Nom_Pro, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cast(Requisiciones.Nom_Pro As Varchar) = '') THEN Proveedores.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Requisiciones.Rif = '') THEN Proveedores.Rif ELSE Requisiciones.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Proveedores.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cast(Requisiciones.Nom_Pro As Varchar) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Requisiciones.Dir_Fis,1, 200) = '') THEN SUBSTRING(Proveedores.Dir_Fis,1, 200) ELSE SUBSTRING(Requisiciones.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cast(Requisiciones.Nom_Pro As Varchar) = '') THEN Proveedores.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Requisiciones.Telefonos = '') THEN Proveedores.Telefonos ELSE Requisiciones.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("           Proveedores.Fax, ")
            loConsulta.AppendLine("           Requisiciones.Nom_Pro                    As  Nom_Gen, ")
            loConsulta.AppendLine("           Requisiciones.Rif                        As  Rif_Gen, ")
            loConsulta.AppendLine("           Requisiciones.Nit                        As  Nit_Gen, ")
            loConsulta.AppendLine("           SPACE(1)                                 As  Dir_Gen, ")
            loConsulta.AppendLine("           SPACE(1)                                 As  Tel_Gen, ")
            loConsulta.AppendLine("           Requisiciones.Documento, ")
            loConsulta.AppendLine("           Requisiciones.Fec_Ini, ")
            loConsulta.AppendLine("           Requisiciones.Fec_Fin, ")
            loConsulta.AppendLine("           Requisiciones.Mon_Bru, ")
            loConsulta.AppendLine("           Requisiciones.Mon_Imp1, ")
            loConsulta.AppendLine("           Requisiciones.Por_Des1, ")
            loConsulta.AppendLine("           Requisiciones.Por_Rec1, ")
            loConsulta.AppendLine("           Requisiciones.Mon_Des1, ")
            loConsulta.AppendLine("           Requisiciones.Mon_Rec1, ")
            loConsulta.AppendLine("           Requisiciones.Dis_Imp, ")
            loConsulta.AppendLine("           Requisiciones.Mon_Net, ")
            loConsulta.AppendLine("           Requisiciones.Cod_For, ")
            loConsulta.AppendLine("           SUBSTRING(Formas_Pagos.Nom_For,1,24)    AS  Nom_For, ")
            loConsulta.AppendLine("           Requisiciones.Cod_Ven, ")
            loConsulta.AppendLine("           Requisiciones.Comentario, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Renglones_Requisiciones.Cod_Art, ")
            loConsulta.AppendLine("           Articulos.Nom_Art, ")
            loConsulta.AppendLine("           Renglones_Requisiciones.Renglon, ")
            loConsulta.AppendLine("           Renglones_Requisiciones.Can_Art1, ")
            loConsulta.AppendLine("           Renglones_Requisiciones.Cod_Uni, ")
            loConsulta.AppendLine("           Renglones_Requisiciones.Por_Des, ")
            loConsulta.AppendLine("           Renglones_Requisiciones.Precio1, ")
            loConsulta.AppendLine("           Renglones_Requisiciones.Mon_Net          As  Neto, ")
            loConsulta.AppendLine("           Renglones_Requisiciones.Por_Imp1         As  Por_Imp, ")
            loConsulta.AppendLine("           Renglones_Requisiciones.Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_Requisiciones.Mon_Imp1         As  Impuesto, ")
            loConsulta.AppendLine("           Sucursales.nom_suc                         AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("           COALESCE(campos_propiedades.val_car, '')   AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.direccion                       AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.telefonos                       AS Telefono_Empresa_Cliente,  ") 
            loConsulta.AppendLine("           Sucursales.Fax                             AS Fax_Empresa_Cliente   ") 
            loConsulta.AppendLine("FROM       Requisiciones ")
            loConsulta.AppendLine("    JOIN   Renglones_Requisiciones ON Requisiciones.Documento  =   Renglones_Requisiciones.Documento")
            loConsulta.AppendLine("    JOIN   Proveedores ON Requisiciones.Cod_Pro  = Proveedores.Cod_Pro")
            loConsulta.AppendLine("    JOIN   Formas_Pagos ON Requisiciones.Cod_For = Formas_Pagos.Cod_For")
            loConsulta.AppendLine("    JOIN   Vendedores ON Requisiciones.Cod_Ven   = Vendedores.Cod_Ven")
            loConsulta.AppendLine("    JOIN   Articulos ON Articulos.Cod_Art        = Renglones_Requisiciones.Cod_Art")
            loConsulta.AppendLine("    JOIN   Sucursales ON Sucursales.cod_suc      = Requisiciones.cod_suc")
            loConsulta.AppendLine("    LEFT JOIN campos_propiedades ON campos_propiedades.cod_reg = Sucursales.cod_suc")
            loConsulta.AppendLine("            AND campos_propiedades.origen = 'Sucursales'")
            loConsulta.AppendLine("            AND campos_propiedades.cod_pro = 'SUC-RIF'")
            loConsulta.AppendLine(" WHERE      " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos
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
                        lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & goServicios.mObtenerFormatoCadena(CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText)) & "%"
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fRequisiciones_Internas_IKP", laDatosReporte)
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text29"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)   
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfRequisiciones_Internas_IKP.ReportSource = loObjetoReporte

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
' RJG: 03/08/13: Codigo inicial
'-------------------------------------------------------------------------------------------'
