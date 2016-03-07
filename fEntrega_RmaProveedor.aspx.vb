'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fEntrega_RmaProveedor"
'-------------------------------------------------------------------------------------------'
Partial Class fEntrega_RmaProveedor
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT      Rma.Cod_Reg                            AS Cod_Reg, ")
            loConsulta.AppendLine("            (CASE WHEN (RTRIM(Rma.Nom_Reg) = '') ")
            loConsulta.AppendLine("                THEN Proveedores.Nom_Pro ")
            loConsulta.AppendLine("                ELSE Rma.Nom_Reg")
            loConsulta.AppendLine("            END)                                    AS Nom_Reg, ")
            loConsulta.AppendLine("            (CASE WHEN (Rma.Rif = '') ")
            loConsulta.AppendLine("                THEN Proveedores.Rif ")
            loConsulta.AppendLine("                ELSE Rma.Rif")
            loConsulta.AppendLine("            END)                                    AS Rif,")
            loConsulta.AppendLine("            Proveedores.Nit                            AS Nit, ")
            loConsulta.AppendLine("            (CASE WHEN (Rma.Dir_Fis = '') ")
            loConsulta.AppendLine("                THEN Proveedores.Dir_Fis ")
            loConsulta.AppendLine("                ELSE Rma.Dir_Fis")
            loConsulta.AppendLine("            END)                                    AS Dir_Fis, ")
            loConsulta.AppendLine("            (CASE WHEN (Rma.Telefonos = '') ")
            loConsulta.AppendLine("                THEN Proveedores.Telefonos ")
            loConsulta.AppendLine("                ELSE Rma.Telefonos")
            loConsulta.AppendLine("            END)                                    AS Telefonos, ")
            loConsulta.AppendLine("            Proveedores.Fax                            AS Fax, ")
            loConsulta.AppendLine("            Proveedores.Generico                       AS Generico, ")
            loConsulta.AppendLine("            Rma.Nom_Reg                             AS Nom_Gen, ")
            loConsulta.AppendLine("            Rma.Rif                                 AS Rif_Gen, ")
            loConsulta.AppendLine("            Rma.Nit                                 AS Nit_Gen, ")
            loConsulta.AppendLine("            Rma.Dir_Fis                             AS Dir_Gen, ")
            loConsulta.AppendLine("            Rma.Telefonos                           AS Tel_Gen, ")
            loConsulta.AppendLine("            Rma.Documento                           AS Documento, ")
            loConsulta.AppendLine("            Rma.Cod_Doc                             AS Cod_Doc, ")
            loConsulta.AppendLine("            Rma.Adicional                           AS Adicional, ")
            loConsulta.AppendLine("            Rma.Fec_Ini                             AS Fec_Ini, ")
            loConsulta.AppendLine("            Rma.Fec_Fin                             AS Fec_Fin, ")
            loConsulta.AppendLine("            Rma.Mon_Bru                             AS Mon_Bru, ")
            loConsulta.AppendLine("            Rma.Mon_Imp1                            AS Mon_Imp1, ")
            loConsulta.AppendLine("            Rma.Por_Imp1                            AS Por_Imp1, ")
            loConsulta.AppendLine("            Rma.Mon_Net                             AS Mon_Net, ")
            loConsulta.AppendLine("            Rma.Por_Des1                            AS Por_Des1, ")
            loConsulta.AppendLine("            Rma.Dis_Imp                             AS Dis_Imp, ")
            loConsulta.AppendLine("            Rma.Mon_Des1                            AS Mon_Des, ")
            loConsulta.AppendLine("            Rma.Por_Rec1                            AS Por_Rec1, ")
            loConsulta.AppendLine("            Rma.Mon_Rec1                            AS Mon_Rec, ")
            loConsulta.AppendLine("            Rma.Cod_For                             AS Cod_For, ")
            loConsulta.AppendLine("            SUBSTRING(Formas_Pagos.Nom_For,1,25)    AS Nom_For, ")
            loConsulta.AppendLine("            Rma.Cod_Ven                             AS Cod_Ven, ")
            loConsulta.AppendLine("            Rma.Comentario                          AS Comentario, ")
            loConsulta.AppendLine("            Vendedores.Nom_Ven                      AS Nom_Ven, ")
            loConsulta.AppendLine("            Renglones_Rma.Cod_Art                   AS Cod_Art, ")
            loConsulta.AppendLine("            CASE WHEN Articulos.Generico = 0 ")
            loConsulta.AppendLine("                THEN Articulos.Nom_Art ")
            loConsulta.AppendLine("		        ELSE Renglones_Rma.Notas END        AS Nom_Art,  ")
            loConsulta.AppendLine("            Renglones_Rma.Renglon                   AS Renglon, ")
            loConsulta.AppendLine("            (CASE WHEN (Renglones_Rma.Cod_Uni2='') ")
            loConsulta.AppendLine("                THEN Renglones_Rma.Can_Art1")
            loConsulta.AppendLine("		        ELSE Renglones_Rma.Can_Art2 END)    AS Can_Art1, ")
            loConsulta.AppendLine("            (CASE WHEN (Renglones_Rma.Cod_Uni2='') ")
            loConsulta.AppendLine("                THEN Renglones_Rma.Cod_Uni")
            loConsulta.AppendLine("		        ELSE Renglones_Rma.Cod_Uni2 END)    AS Cod_Uni, ")
            loConsulta.AppendLine("            (CASE WHEN (Renglones_Rma.Cod_Uni2='') ")
            loConsulta.AppendLine("                THEN Renglones_Rma.Precio1")
            loConsulta.AppendLine("		        ELSE Renglones_Rma.Precio1*Renglones_Rma.Can_Uni2 END) AS Precio1, ")
            loConsulta.AppendLine("            Renglones_Rma.Mon_Net                   AS Neto, ")
            loConsulta.AppendLine("            Renglones_Rma.Por_Imp1                  AS Por_Imp, ")
            loConsulta.AppendLine("            Renglones_Rma.Cod_Imp                   AS Cod_Imp, ")
            loConsulta.AppendLine("            Renglones_Rma.Mon_Imp1                  AS Impuesto ")
            loConsulta.AppendLine("FROM        Rma ")
            loConsulta.AppendLine("    JOIN    Renglones_Rma")
            loConsulta.AppendLine("        ON  Rma.Documento   =   Renglones_Rma.Documento")
            loConsulta.AppendLine("        AND Rma.Cod_Doc     =   Renglones_Rma.Cod_Doc")
            loConsulta.AppendLine("        AND Rma.Adicional   =   Renglones_Rma.Adicional")
            loConsulta.AppendLine("    JOIN    Proveedores")
            loConsulta.AppendLine("        ON  Rma.Cod_Reg    =   Proveedores.Cod_Pro")
            loConsulta.AppendLine("    JOIN    Formas_Pagos")
            loConsulta.AppendLine("        ON  Rma.Cod_For    =   Formas_Pagos.Cod_For")
            loConsulta.AppendLine("    JOIN    Vendedores ")
            loConsulta.AppendLine("        ON  Rma.Cod_Ven    =   Vendedores.Cod_Ven")
            loConsulta.AppendLine("    JOIN    Articulos ")
            loConsulta.AppendLine("        ON  Articulos.Cod_Art   =   Renglones_Rma.Cod_Art")
            loConsulta.AppendLine("WHERE       " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            
            Dim loServicios As New cusDatos.goDatos()
            
            'Me.mEscribirConsulta(loConsulta.ToString())

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

  
            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fEntrega_RmaProveedor", laDatosReporte)
			lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".",",")
            CType(loObjetoReporte.ReportDefinition.ReportObjects("txtRpt_Porcentajes_Impuesto"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfEntrega_RmaProveedor.ReportSource = loObjetoReporte

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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 06/02/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
