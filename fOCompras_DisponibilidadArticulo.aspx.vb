'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fOCompras_DisponibilidadArticulo"
'-------------------------------------------------------------------------------------------'
Partial Class fOCompras_DisponibilidadArticulo
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Ordenes_Compras.Documento                           AS Documento,")
            loConsulta.AppendLine("            Ordenes_Compras.Fec_Ini                             AS Fec_Ini,")
            loConsulta.AppendLine("            Ordenes_Compras.Fec_Fin                             AS Fec_Fin,")
            loConsulta.AppendLine("            Ordenes_Compras.Cod_Pro                             AS Cod_Pro,")
            loConsulta.AppendLine("            Ordenes_Compras.Cod_Ven                             AS Cod_Ven,")
            loConsulta.AppendLine("            Vendedores.Nom_Ven                                   AS Nom_Ven,")
            loConsulta.AppendLine("            Ordenes_Compras.Comentario                          AS Comentario,")
            loConsulta.AppendLine("            (CASE WHEN Ordenes_Compras.Nom_Pro>'' ")
            loConsulta.AppendLine("                THEN Ordenes_Compras.Nom_Pro")
            loConsulta.AppendLine("                ELSE Proveedores.Nom_Pro END)                   AS Nom_Pro,")
            loConsulta.AppendLine("            (CASE WHEN Ordenes_Compras.Rif>'' ")
            loConsulta.AppendLine("                THEN Ordenes_Compras.Rif")
            loConsulta.AppendLine("                ELSE Proveedores.Rif END)                       AS Rif,")
            loConsulta.AppendLine("            (CASE WHEN Ordenes_Compras.Dir_Fis>'' ")
            loConsulta.AppendLine("                THEN Ordenes_Compras.Dir_Fis")
            loConsulta.AppendLine("                ELSE Proveedores.Dir_Fis END)                   AS Dir_Fis,")
            loConsulta.AppendLine("            (CASE WHEN Ordenes_Compras.Telefonos>'' ")
            loConsulta.AppendLine("                THEN Ordenes_Compras.Telefonos")
            loConsulta.AppendLine("                ELSE Proveedores.Telefonos END)                 AS Telefonos,")
            loConsulta.AppendLine("            Proveedores.Fax                                     AS Fax,")
            loConsulta.AppendLine("            Renglones.Renglon                                   AS Renglon,")
            loConsulta.AppendLine("            Renglones.Cod_Art                                   AS Cod_Art,")
            loConsulta.AppendLine("            Renglones.Cod_Uni                                   AS Cod_Uni,")
            loConsulta.AppendLine("            Renglones.Can_Art1                                  AS Can_Art1,")
            loConsulta.AppendLine("            Renglones.Nom_Art                                   AS Nom_Art,")
            loConsulta.AppendLine("            COALESCE(Proveedores_Disponibles.Cod_Pro, '[N/D]')  AS Cod_Pro_Dis, ")
            loConsulta.AppendLine("            COALESCE(Proveedores_Disponibles.Nom_Pro, ")
            loConsulta.AppendLine("                '[Artículos no disponibles en proveedores]')    AS Nom_Pro_Dis, ")
            loConsulta.AppendLine("            Renglones.Items_Necesarios                          AS Items_Necesarios,")
            loConsulta.AppendLine("            COUNT(Renglones.Cod_Art) ")
            loConsulta.AppendLine("                OVER( PARTITION BY COALESCE(Proveedores_Disponibles.Cod_Pro, '')) AS Items_Disponibles ")
            loConsulta.AppendLine("FROM        Ordenes_Compras")
            loConsulta.AppendLine("        JOIN (  SELECT  Renglones_oCompras.Documento                            AS Documento,")
            loConsulta.AppendLine("                        Min(Renglones_oCompras.Renglon)                         AS Renglon,")
            loConsulta.AppendLine("                        Renglones_oCompras.Cod_Art                              AS Cod_Art, ")
            loConsulta.AppendLine("                        Renglones_oCompras.Cod_Uni                              AS Cod_Uni,")
            loConsulta.AppendLine("                        SUM(Renglones_oCompras.Can_Art1)                        AS Can_Art1,")
            loConsulta.AppendLine("                        Renglones_oCompras.Notas                                AS Nom_Art,")
            loConsulta.AppendLine("                        COUNT(Renglones_oCompras.Cod_Art) ")
            loConsulta.AppendLine("                            OVER( PARTITION BY Renglones_oCompras.Documento)    AS Items_Necesarios")
            loConsulta.AppendLine("                        ")
            loConsulta.AppendLine("                FROM    Renglones_oCompras  ")
            loConsulta.AppendLine("                GROUP BY Renglones_oCompras.Documento,")
            loConsulta.AppendLine("                        Renglones_oCompras.Cod_Art,")
            loConsulta.AppendLine("                        Renglones_oCompras.Cod_Uni,")
            loConsulta.AppendLine("                        Renglones_oCompras.Notas")
            loConsulta.AppendLine("                ) Renglones ")
            loConsulta.AppendLine("            ON  Renglones.Documento = Ordenes_Compras.Documento ")
            loConsulta.AppendLine("        LEFT JOIN ( SELECT      Renglones_lCompras.Cod_Art,")
            loConsulta.AppendLine("                                Libres_Compras.Cod_Pro")
            loConsulta.AppendLine("                    FROM        Renglones_lCompras")
            loConsulta.AppendLine("                        JOIN    Libres_Compras ")
            loConsulta.AppendLine("                            ON  Libres_Compras.Documento = Renglones_lCompras.Documento")
            loConsulta.AppendLine("                            AND Libres_Compras.Status <> 'Anulado'")
            loConsulta.AppendLine("                    GROUP BY    Renglones_lCompras.Cod_Art, ")
            loConsulta.AppendLine("                                Libres_Compras.Cod_Pro")
            loConsulta.AppendLine("                    ) Disponibles ON  Disponibles.Cod_Art = Renglones.Cod_Art")
            loConsulta.AppendLine("        LEFT JOIN Proveedores Proveedores_Disponibles ON Proveedores_Disponibles.Cod_Pro = Disponibles.Cod_Pro")
            loConsulta.AppendLine("            AND Proveedores_Disponibles.Status = 'A'")
            loConsulta.AppendLine("        JOIN Proveedores ON Proveedores.Cod_Pro = Ordenes_Compras.Cod_Pro")
            loConsulta.AppendLine("        JOIN Vendedores ON Vendedores.Cod_Ven = Ordenes_Compras.Cod_Ven")
            loConsulta.AppendLine("WHERE       " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("ORDER BY    Renglon ASC, COALESCE(Proveedores_Disponibles.Cod_Pro, '') ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

			'ME.mEscribirConsulta(loConsulta.ToString())
			 
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")
            
            
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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fOCompras_DisponibilidadArticulo", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfOCompras_DisponibilidadArticulo.ReportSource = loObjetoReporte

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
' RJG: 04/09/15: Código Inicial, a partir del Formato de Orden de Compra.
'-------------------------------------------------------------------------------------------'
