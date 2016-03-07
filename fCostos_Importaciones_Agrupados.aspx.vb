'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCostos_Importaciones_Agrupados"
'-------------------------------------------------------------------------------------------'
Partial Class fCostos_Importaciones_Agrupados
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()
            
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT      Importaciones.Documento                             AS Documento,")
			loConsulta.AppendLine("            Importaciones.Fec_Ini                               AS Fec_Ini,")
			loConsulta.AppendLine("            Importaciones.Status                                AS Status,")
			loConsulta.AppendLine("            Importaciones.Comentario                            AS Comentario,")
			loConsulta.AppendLine("            Importaciones.Expediente                            AS Expediente,")
			loConsulta.AppendLine("            Importaciones.Factura                               AS Factura,")
			loConsulta.AppendLine("            Importaciones.Age_Adu                               AS Age_Adu,")
			loConsulta.AppendLine("            COALESCE(Agente_Aduanal.Nom_Pro, '[NO DISPONIBLE]') AS Age_Adu_Nombre,")
			loConsulta.AppendLine("            Importaciones.Exp_Adu                               AS Exp_Adu,")
			loConsulta.AppendLine("            Importaciones.Exp_Alm                               AS Exp_Alm,")
			loConsulta.AppendLine("            Importaciones.Total_Contenedores                    AS Total_Contenedores,")
			loConsulta.AppendLine("            Importaciones.Numero_BL                             AS Numero_BL,")
			loConsulta.AppendLine("            Importaciones.Modalidad                             AS Modalidad,")
			loConsulta.AppendLine("            Importaciones.Obt_Rec                               AS Obt_Rec,")
			loConsulta.AppendLine("            Importaciones.Tas_Emi                               AS Tas_Emi,")
			loConsulta.AppendLine("            Detalles_Importaciones.Concepto                     AS Concepto,")
			loConsulta.AppendLine("            Detalles_Importaciones.Cod_Gas                      AS Cod_Gas,")
			loConsulta.AppendLine("            Detalles_Importaciones.Nom_Gas                      AS Nom_Gas,")
			loConsulta.AppendLine("            Detalles_Importaciones.Mon_Gas                      AS Mon_Gas")
			loConsulta.AppendLine("FROM        Importaciones")
			loConsulta.AppendLine("    JOIN    Detalles_Importaciones ")
			loConsulta.AppendLine("        ON  Detalles_Importaciones.Grupo = 'Gastos Fijos'")
			loConsulta.AppendLine("        AND Detalles_Importaciones.Origen = Importaciones.Origen")
			loConsulta.AppendLine("        AND Detalles_Importaciones.Adicional = Importaciones.Adicional")
			loConsulta.AppendLine("        AND Detalles_Importaciones.Documento = Importaciones.Documento")
			loConsulta.AppendLine("        AND Detalles_Importaciones.Afe_Cos = 1")
			loConsulta.AppendLine("    LEFT JOIN Proveedores Agente_Aduanal ")
			loConsulta.AppendLine("        ON  Agente_Aduanal.Cod_Pro = Importaciones.Age_Adu")
			loConsulta.AppendLine("WHERE       " & cusAplicacion.goFormatos.pcCondicionPrincipal)
			loConsulta.AppendLine("ORDER BY    Detalles_Importaciones.Concepto ASC, Detalles_Importaciones.Renglon ASC")
			loConsulta.AppendLine("")
			
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
			
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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCostos_Importaciones_Agrupados", laDatosReporte)
            
            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCostos_Importaciones_Agrupados.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 14/12/14: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
