'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fGuias_Clientes_IKP"
'-------------------------------------------------------------------------------------------'
Partial Class fGuias_Clientes_IKP

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" SELECT	Guias.Cod_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Guias.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Guias.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Guias.Nom_Cli END) END) AS  Nom_Cli, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Guias.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Guias.Rif = '') THEN Clientes.Rif ELSE Guias.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("           Clientes.Nit, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Guias.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Guias.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Guias.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("           (CASE WHEN (Clientes.Generico = 0 AND Guias.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Guias.Telefonos = '') THEN Clientes.Telefonos ELSE Guias.Telefonos END) END) AS  Telefonos, ")
            
            loConsulta.AppendLine("           Clientes.Fax, ")
            loConsulta.AppendLine("           Guias.Nom_Cli                    As  Nom_Gen, ")
            loConsulta.AppendLine("           Guias.Rif                        As  Rif_Gen, ")
            loConsulta.AppendLine("           Guias.Nit                        As  Nit_Gen, ")
            loConsulta.AppendLine("           Guias.Dir_Fis                    As  Dir_Gen, ")
            loConsulta.AppendLine("           Guias.Telefonos                  As  Tel_Gen, ")
            loConsulta.AppendLine("           Guias.Documento, ")
            loConsulta.AppendLine("           Guias.Fec_Ini, ")
            loConsulta.AppendLine("           Guias.Fec_Fin, ")
            loConsulta.AppendLine("           Guias.Mon_Bru, ")
            loConsulta.AppendLine("           Guias.Por_Des1, ")
            loConsulta.AppendLine("           Guias.Por_Rec1, ")
            loConsulta.AppendLine("           Guias.Mon_Des1, ")
            loConsulta.AppendLine("           Guias.Mon_Rec1, ")
            loConsulta.AppendLine("           Guias.Mon_Imp1, ")
            loConsulta.AppendLine("           Guias.Mon_Net, ")
            loConsulta.AppendLine("           Guias.Cod_For, ")
            loConsulta.AppendLine("           Guias.Dis_Imp, ")
            loConsulta.AppendLine("           Formas_Pagos.Nom_For, ")
            loConsulta.AppendLine("           Guias.Cod_Ven, ")
            loConsulta.AppendLine("           Guias.Comentario, ")
            loConsulta.AppendLine("           Vendedores.Nom_Ven, ")
            loConsulta.AppendLine("           Renglones_Guias.Cod_Art, ")
            loConsulta.AppendLine("           Articulos.Nom_Art, ")
            loConsulta.AppendLine("           Renglones_Guias.Renglon, ")
            loConsulta.AppendLine("           Renglones_Guias.Por_des, ")
            loConsulta.AppendLine("           Renglones_Guias.Can_Art1, ")
            loConsulta.AppendLine("           Renglones_Guias.Cod_Uni, ")
            loConsulta.AppendLine("           Renglones_Guias.Precio1, ")
            loConsulta.AppendLine("           Renglones_Guias.Mon_Net          As  Neto, ")
            loConsulta.AppendLine("           Renglones_Guias.Por_Imp1         As  Por_Imp, ")
            loConsulta.AppendLine("           Renglones_Guias.Cod_Imp, ")
            loConsulta.AppendLine("           Renglones_Guias.Mon_Imp1         As  Impuesto, ")
            loConsulta.AppendLine("           Sucursales.nom_suc                       AS Nombre_Empresa_Cliente,  ")
            loConsulta.AppendLine("           COALESCE(campos_propiedades.val_car, '') AS Rif_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.direccion                     AS Direccion_Empresa_Cliente,  ")
            loConsulta.AppendLine("           Sucursales.telefonos                     AS Telefono_Empresa_Cliente   ")
            loConsulta.AppendLine(" FROM      Guias")
            loConsulta.AppendLine("    JOIN   Renglones_Guias ON Guias.Documento=   Renglones_Guias.Documento")
            loConsulta.AppendLine("    JOIN   Clientes ON Guias.Cod_Cli         =   Clientes.Cod_Cli")
            loConsulta.AppendLine("    JOIN   Formas_Pagos ON Guias.Cod_For     =   Formas_Pagos.Cod_For")
            loConsulta.AppendLine("    JOIN   Vendedores ON Guias.Cod_Ven       =   Vendedores.Cod_Ven")
            loConsulta.AppendLine("    JOIN   Articulos ON Articulos.Cod_Art    =   Renglones_Guias.Cod_Art")
            loConsulta.AppendLine("    JOIN   Sucursales ON Sucursales.cod_suc  = Guias.cod_suc")
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

                'En cada renglón lee el contenido de la distribució de impuestos
                For Each loImpuesto As System.Xml.XmlNode In loImpuestos.SelectNodes("impuestos/impuesto")
                    If lnNumeroFila = laDatosReporte.Tables(0).Rows.Count - 1 Then
                        lcPorcentajesImpueto = lcPorcentajesImpueto & ", " & CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText) & "%"
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
            
            
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fGuias_Clientes_IKP", laDatosReporte)
            
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text24"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfGuias_Clientes_IKP.ReportSource = loObjetoReporte

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
' JJD: 13/12/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 24/12/09: Se le incluyo los montos de descuentos y recargos.
'-------------------------------------------------------------------------------------------'
' CMS: 18/03/10: Se aplicaron los metodos carga de imagen y validacion de registro cero
'-------------------------------------------------------------------------------------------'
' CMS: 21/04/10: Se ajusto para que tome el cliente generico
'-------------------------------------------------------------------------------------------'