'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fFacturas_Ventas_MULTIBIZ"
'-------------------------------------------------------------------------------------------'
Partial Class fFacturas_Ventas_MULTIBIZ
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Facturas.Cod_Cli                            AS Cod_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Facturas.Nom_Cli = '') THEN Clientes.Nom_Cli ELSE Facturas.Nom_Cli END) END) AS  Nom_Cli, ")
            loConsulta.AppendLine("            (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Rif ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Facturas.Rif = '') THEN Clientes.Rif ELSE Facturas.Rif END) END) AS  Rif, ")
            loConsulta.AppendLine("            Clientes.Nit                                AS Nit, ")
            loConsulta.AppendLine("            (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (SUBSTRING(Facturas.Dir_Fis,1, 200) = '') THEN SUBSTRING(Clientes.Dir_Fis,1, 200) ELSE SUBSTRING(Facturas.Dir_Fis,1, 200) END) END) AS  Dir_Fis, ")
            loConsulta.AppendLine("            (CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '') THEN Clientes.Telefonos ELSE ")
            loConsulta.AppendLine("               (CASE WHEN (Facturas.Telefonos = '') THEN Clientes.Telefonos ELSE Facturas.Telefonos END) END) AS  Telefonos, ")
            loConsulta.AppendLine("            Clientes.Fax                                AS Fax, ")
            loConsulta.AppendLine("            Clientes.Generico                           AS Generico, ")
            loConsulta.AppendLine("            Facturas.Nom_Cli                            AS Nom_Gen, ")
            loConsulta.AppendLine("            Facturas.Rif                                AS Rif_Gen, ")
            loConsulta.AppendLine("            Facturas.Nit                                AS Nit_Gen, ")
            loConsulta.AppendLine("            Facturas.Dir_Fis                            AS Dir_Gen, ")
            loConsulta.AppendLine("            Facturas.Telefonos                          AS Tel_Gen, ")
            loConsulta.AppendLine("            Facturas.Documento                          AS Documento, ")
            loConsulta.AppendLine("            Facturas.Fec_Ini                            AS Fec_Ini, ")
            loConsulta.AppendLine("            Facturas.Fec_Fin                            AS Fec_Fin, ")
            loConsulta.AppendLine("            Facturas.Mon_Bru                            AS Mon_Bru, ")
            loConsulta.AppendLine("            Facturas.Mon_Imp1                           AS Mon_Imp1, ")
            loConsulta.AppendLine("            Facturas.Por_Imp1                           AS Por_Imp1, ")
            loConsulta.AppendLine("            Facturas.Mon_Net                            AS Mon_Net, ")
            loConsulta.AppendLine("            Facturas.Por_Des1                           AS Por_Des1, ")
            loConsulta.AppendLine("            Facturas.Dis_Imp                            AS Dis_Imp, ")
            loConsulta.AppendLine("            Facturas.Mon_Des1                           AS Mon_Des, ")
            loConsulta.AppendLine("            Facturas.Por_Rec1                           AS Por_Rec1, ")
            loConsulta.AppendLine("            Facturas.Mon_Rec1                           AS Mon_Rec, ")
            loConsulta.AppendLine("            Facturas.Cod_For                            AS Cod_For, ")
            loConsulta.AppendLine("            Formas_Pagos.Nom_For                        AS Nom_For, ")
            loConsulta.AppendLine("            Transportes.Nom_Tra                         AS Nom_Tra, ")
            loConsulta.AppendLine("            Facturas.Cod_Ven                            AS Cod_Ven, ")
            loConsulta.AppendLine("            Facturas.Comentario                         AS Comentario, ")
            loConsulta.AppendLine("            Vendedores.Nom_Ven                          AS Nom_Ven, ")
            loConsulta.AppendLine("            Renglones_Facturas.Cod_Art                  AS Cod_Art, ")
            loConsulta.AppendLine("            Articulos.Garantia                          AS Garantia, ")
            loConsulta.AppendLine("            CASE WHEN Articulos.Generico = 0 ")
            loConsulta.AppendLine("                THEN Articulos.Nom_Art ")
            loConsulta.AppendLine("		        ELSE Renglones_Facturas.Notas END           AS Nom_Art,  ")
            loConsulta.AppendLine("            Renglones_Facturas.Renglon                  AS Renglon, ")
            loConsulta.AppendLine("            (CASE WHEN (Renglones_Facturas.Cod_Uni2='') ")
            loConsulta.AppendLine("                THEN Renglones_Facturas.Can_Art1")
            loConsulta.AppendLine("		        ELSE Renglones_Facturas.Can_Art2 END)       AS Can_Art1, ")
            loConsulta.AppendLine("            (CASE WHEN (Renglones_Facturas.Cod_Uni2='') ")
            loConsulta.AppendLine("                THEN Renglones_Facturas.Cod_Uni")
            loConsulta.AppendLine("		        ELSE Renglones_Facturas.Cod_Uni2 END)       AS Cod_Uni, ")
            loConsulta.AppendLine("            (CASE WHEN (Renglones_Facturas.Cod_Uni2='') ")
            loConsulta.AppendLine("                THEN Renglones_Facturas.Precio1")
            loConsulta.AppendLine("		        ELSE Renglones_Facturas.Precio1*Renglones_Facturas.Can_Uni2 END) AS Precio1,")
            loConsulta.AppendLine("            Renglones_Facturas.Mon_Net                  AS Neto, ")
            loConsulta.AppendLine("            Renglones_Facturas.Por_Imp1                 AS Por_Imp, ")
            loConsulta.AppendLine("            Renglones_Facturas.Cod_Imp                  AS Cod_Imp, ")
            loConsulta.AppendLine("            Renglones_Facturas.Mon_Imp1                 AS Impuesto, ")
            loConsulta.AppendLine("            Facturas.Mon_Exe                                                                   AS Exento,")
            loConsulta.AppendLine("            ROUND(Facturas.Mon_Bru - Facturas.Mon_Exe, 2)                                      AS Grabable,")
            loConsulta.AppendLine("            ROUND(Facturas.Mon_Des1*Facturas.Mon_Exe/Facturas.Mon_Bru, 2)                      AS Descuento_Exentos,")
            loConsulta.AppendLine("            ROUND(Facturas.Mon_Rec1*Facturas.Mon_Exe/Facturas.Mon_Bru, 2)                      AS Recargo_Exentos,")
            loConsulta.AppendLine("            ROUND(Facturas.Mon_Des1*(Facturas.Mon_Bru - Facturas.Mon_Exe)/Facturas.Mon_Bru, 2) AS Descuento_Grabable,")
            loConsulta.AppendLine("            ROUND(Facturas.Mon_Rec1*(Facturas.Mon_Bru - Facturas.Mon_Exe)/Facturas.Mon_Bru, 2) AS Recargo_Grabable")
            loConsulta.AppendLine("FROM        Facturas ")
            loConsulta.AppendLine("    JOIN    Renglones_Facturas")
            loConsulta.AppendLine("        ON  Facturas.Documento  =   Renglones_Facturas.Documento")
            loConsulta.AppendLine("    JOIN    Clientes")
            loConsulta.AppendLine("        ON  Facturas.Cod_Cli    =   Clientes.Cod_Cli")
            loConsulta.AppendLine("    JOIN    Formas_Pagos")
            loConsulta.AppendLine("        ON  Facturas.Cod_For    =   Formas_Pagos.Cod_For")
            loConsulta.AppendLine("    JOIN    Transportes")
            loConsulta.AppendLine("        ON  Facturas.Cod_Tra    =   Transportes.Cod_Tra")
            loConsulta.AppendLine("    JOIN    Vendedores ")
            loConsulta.AppendLine("        ON  Facturas.Cod_Ven    =   Vendedores.Cod_Ven")
            loConsulta.AppendLine("    JOIN    Articulos ")
            loConsulta.AppendLine("        ON  Articulos.Cod_Art   =   Renglones_Facturas.Cod_Art")
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
									lcPorcentajesImpueto = lcPorcentajesImpueto & ", " _
									    & goServicios.mObtenerFormatoCadena(CDec(loImpuesto.SelectSingleNode("porcentaje").InnerText), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 0) _
									    & "%"
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fFacturas_Ventas_MULTIBIZ", laDatosReporte)
			lcPorcentajesImpueto = lcPorcentajesImpueto.Replace(".",",")
            CType(loObjetoReporte.ReportDefinition.ReportObjects("Text1"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = lcPorcentajesImpueto.ToString

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfFacturas_Ventas_MULTIBIZ.ReportSource = loObjetoReporte

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
' RJG: 19/05/15: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
