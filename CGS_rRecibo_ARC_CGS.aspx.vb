'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rRecibo_ARC_CGS"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rRecibo_ARC_CGS
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

            Dim lcSoloConReteniciones As String = CStr(cusAplicacion.goReportes.paParametrosFinales(4)).Trim().ToUpper()

            Dim lcConceptosDeRetencion As String = "R005,R006,R305,R405"
            Dim lcConceptosAdicionales As String = "A200,A304,A305,A402,A404,A405,A406,A407"
            Dim lcConceptosAExcluir As String = "A300,A301,A302,A303"

            lcConceptosDeRetencion  = goServicios.mObtenerListaFormatoSQL(lcConceptosDeRetencion)
            lcConceptosAdicionales  = goServicios.mObtenerListaFormatoSQL(lcConceptosAdicionales)
            lcConceptosAExcluir = goServicios.mObtenerListaFormatoSQL(lcConceptosAExcluir)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
                
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE @lcDcto_Desde AS VARCHAR(10) = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @lcDcto_Hasta AS VARCHAR(10) = " & lcParametro0Hasta)
            loConsulta.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro1Desde)
            loConsulta.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro1Hasta)
            loConsulta.AppendLine("DECLARE @lcCodTra_Desde AS VARCHAR(9) = " & lcParametro2Desde)
            loConsulta.AppendLine("DECLARE @lcCodTra_Hasta AS VARCHAR(9) = " & lcParametro2Hasta)
            loConsulta.AppendLine("DECLARE @lcCodCon_Desde AS VARCHAR(2) = " & lcParametro3Desde)
            loConsulta.AppendLine("DECLARE @lcCodCon_Hasta AS VARCHAR(2) = " & lcParametro3Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT     *")
            loConsulta.AppendLine("FROM (")
            loConsulta.AppendLine("SELECT      Trabajadores.Cod_Tra                     AS Codigo,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                     AS Nombre,")
            loConsulta.AppendLine("            Trabajadores.Cedula                      AS Cedula,")
            loConsulta.AppendLine("            Trabajadores.Rif                         AS Rif,")
            loConsulta.AppendLine("            YEAR(Recibos.Fecha )                     AS Anio,")
            loConsulta.AppendLine("            MONTH(Recibos.Fecha)                     AS Mes,")
            loConsulta.AppendLine("            SUM(Retenciones.Monto)                   AS Monto,")
            loConsulta.AppendLine("            SUM(Retenciones.Retenido)                AS Retenido,")
            loConsulta.AppendLine("            ROUND(SUM(Retenciones.Retenido)*100")
            loConsulta.AppendLine("				/SUM(Retenciones.Monto),2)              AS Porcentaje,")
            loConsulta.AppendLine("            CAST(@ldFecha_Desde AS DATE)             AS Fecha_Desde,")
            loConsulta.AppendLine("            CAST(@ldFecha_Hasta AS DATE)             AS Fecha_Hasta,")
            loConsulta.AppendLine("			   SUM(SUM(Retenciones.Retenido)) OVER (PARTITION BY Trabajadores.Cod_Tra)  AS Total_Retenido")
            loConsulta.AppendLine("FROM        Recibos")
            loConsulta.AppendLine("    JOIN    (   SELECT  RR.Documento													AS Documento,")
            loConsulta.AppendLine("						SUM(CASE WHEN RR.Cod_Con IN (" & lcConceptosDeRetencion & ")")
            loConsulta.AppendLine("							THEN RR.Mon_Net ")
            loConsulta.AppendLine("							ELSE 0 END)													AS Retenido,")
            loConsulta.AppendLine("						MAX(CASE WHEN RR.Cod_Con IN (" & lcConceptosDeRetencion & ")")
            loConsulta.AppendLine("							THEN CAST(LEFT(RR.Val_Car,4) AS DECIMAL(28,10))  ")
            loConsulta.AppendLine("							ELSE 0 END)													AS Porcentaje,")
            loConsulta.AppendLine("                     SUM(ROUND(")
            loConsulta.AppendLine("							CASE WHEN RR.Cod_Con IN (" & lcConceptosDeRetencion & ")")
            loConsulta.AppendLine("								AND RR.Val_Num>0 ")
            loConsulta.AppendLine("                            THEN RR.Val_Num")
            loConsulta.AppendLine("                            ELSE ")
            loConsulta.AppendLine("								CASE WHEN RR.Cod_Con IN (" & lcConceptosDeRetencion & ")								")
            loConsulta.AppendLine("								THEN 0")
            loConsulta.AppendLine("								ELSE (CASE WHEN CN.Tipo = 'Asignacion' ")
            loConsulta.AppendLine("										THEN RR.Mon_Net ELSE -RR.Mon_Net END) END")
            loConsulta.AppendLine("						END, 2))														AS Monto ")
            loConsulta.AppendLine("                FROM    Renglones_Recibos RR")
            loConsulta.AppendLine("					LEFT JOIN Conceptos_Nomina CN ON CN.Cod_Con = RR.Cod_Con")
            loConsulta.AppendLine("                WHERE   (RR.Cod_Con IN (" & lcConceptosDeRetencion & ")")
            loConsulta.AppendLine("					OR  RR.Cod_Con IN (" & lcConceptosAdicionales & ")")
            loConsulta.AppendLine("					OR  CN.Acumulados = 1) ")
            loConsulta.AppendLine("					AND RR.Cod_Con NOT IN (" & lcConceptosAExcluir & ")")
            loConsulta.AppendLine("                GROUP BY RR.Documento ")
            'loConsulta.AppendLine("				HAVING	SUM(CASE WHEN RR.Cod_Con IN (" & lcConceptosDeRetencion & ") THEN RR.Mon_Net ELSE 0 END) > 0")
            loConsulta.AppendLine("        ) Retenciones")
            loConsulta.AppendLine("        ON  Retenciones.Documento = Recibos.Documento")
            loConsulta.AppendLine("    JOIN    Trabajadores ON Trabajadores.Cod_Tra = Recibos.Cod_Tra")
            loConsulta.AppendLine("WHERE      ")
            loConsulta.AppendLine("         Recibos.Documento				BETWEEN @lcDcto_Desde AND @lcDcto_Hasta")
            loConsulta.AppendLine(" 	AND Recibos.Fecha	            	BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loConsulta.AppendLine(" 	AND Recibos.Cod_Tra				    BETWEEN @lcCodTra_Desde AND @lcCodTra_Hasta")
            loConsulta.AppendLine(" 	AND Recibos.Cod_Con				    BETWEEN @lcCodCon_Desde AND @lcCodCon_Hasta")
            loConsulta.AppendLine(" 	AND Recibos.Status				    = 'Confirmado'")
            loConsulta.AppendLine("GROUP BY	YEAR(Recibos.Fecha ),")
            loConsulta.AppendLine("			MONTH(Recibos.Fecha),")
            loConsulta.AppendLine("			Trabajadores.Cod_Tra, ")
            loConsulta.AppendLine("			Trabajadores.Nom_Tra, ")
            loConsulta.AppendLine("			Trabajadores.Cedula,  ")
            loConsulta.AppendLine("			Trabajadores.Rif")
            loConsulta.AppendLine(") Resumen")
            If (lcSoloConReteniciones = "SI") Then 
                loConsulta.AppendLine("WHERE     Total_Retenido >0")
            End If
            loConsulta.AppendLine("ORDER BY    " & lcOrdenamiento & ", Anio, Mes")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos
            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rRecibo_ARC_CGS", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)


            'Dim myPrintOptions As CrystalDecisions.CrystalReports.Engine.PrintOptions = loObjetoReporte.PrintOptions



            'Dim ps As New System.Drawing.Printing.PrinterSettings()
            'Dim loTamaño As New Drawing.Printing.PaperSize("MediaCarta", 850, 550) '8,5x5,5 pulgadas 

            'ps.DefaultPageSettings.PaperSize = loTamaño
            'Dim pags As New System.Drawing.Printing.PageSettings(ps)

            ''ps.DefaultPageSettings.PaperSize.Width = 8.5 * 100
            ''ps.DefaultPageSettings.PaperSize.Height = 5.5 * 100

            'loObjetoReporte.PrintOptions.CopyFrom(ps, pags)
            'loObjetoReporte.PrintOptions.PaperSize = CrystalDecisions.Shared.PaperSize.DefaultPaperSize '(ps, ps.DefaultPageSettings)


            Me.crvCGS_rRecibo_ARC_CGS.ReportSource = loObjetoReporte


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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 08/01/16: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
