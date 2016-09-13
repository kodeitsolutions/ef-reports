'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTotales_porConceptos"
'-------------------------------------------------------------------------------------------'
Partial Class rTotales_porConceptos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            'Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Conceptos_Nomina.Cod_Con                AS Cod_Con,")
            loConsulta.AppendLine("            Conceptos_Nomina.Nom_Con                AS Nom_Con,")
            loConsulta.AppendLine("            Conceptos_Nomina.tipo                   AS tipo,")
            loConsulta.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loConsulta.AppendLine("                WHEN 'Asignacion'")
            loConsulta.AppendLine("                THEN Renglones_Recibos.mon_net")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)                                   AS Asignacion,")
            loConsulta.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loConsulta.AppendLine("                WHEN 'Deduccion'")
            loConsulta.AppendLine("                THEN - Renglones_Recibos.mon_net")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)                                   AS Deduccion,")
            loConsulta.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loConsulta.AppendLine("                WHEN 'Retencion'")
            loConsulta.AppendLine("                THEN - Renglones_Recibos.mon_net")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)                                   AS Retencion,")
            loConsulta.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loConsulta.AppendLine("                WHEN 'Otro'")
            loConsulta.AppendLine("                THEN Renglones_Recibos.mon_net")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)                                   AS Otro")
            loConsulta.AppendLine("FROM        Renglones_Recibos")
            loConsulta.AppendLine("    JOIN    Recibos ")
            loConsulta.AppendLine("        ON  Recibos.Documento = Renglones_Recibos.Documento")
            loConsulta.AppendLine("    JOIN    Conceptos_Nomina ")
            loConsulta.AppendLine("        ON  Conceptos_Nomina.cod_con = Renglones_Recibos.cod_con ")
            loConsulta.AppendLine("    JOIN    Trabajadores ")
            loConsulta.AppendLine("        ON  Trabajadores.cod_tra = Recibos.cod_tra ")
            loConsulta.AppendLine("WHERE		 Conceptos_Nomina.Cod_Con BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loConsulta.AppendLine("        AND Conceptos_Nomina.Tipo IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine("        AND Recibos.Cod_Con BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Dep BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Car BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Tra BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loConsulta.AppendLine("        AND Conceptos_Nomina.Cod_Cla BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Suc BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loConsulta.AppendLine("        AND Recibos.Cod_Mon BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta)
            loConsulta.AppendLine("        AND Recibos.Status IN ('Confirmado', 'Procesado')")
            loConsulta.AppendLine("GROUP BY    conceptos_Nomina.Cod_Con,")
            loConsulta.AppendLine("            conceptos_Nomina.Nom_Con,")
            loConsulta.AppendLine("            conceptos_Nomina.Tipo")
            loConsulta.AppendLine("ORDER BY    (CASE Conceptos_Nomina.tipo")
            loConsulta.AppendLine("                WHEN 'Asignacion' THEN 0")
            loConsulta.AppendLine("                WHEN 'Deduccion' THEN 1")
            loConsulta.AppendLine("                WHEN 'Retencion' THEN 1")
            loConsulta.AppendLine("                ELSE 2")
            loConsulta.AppendLine("             END) ASC, cod_con ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTotales_porConceptos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTotales_porConceptos.ReportSource = loObjetoReporte

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
' RJG: 20/12/13: Codigo inicial, a partir de rTotales_Conceptos.							'
'-------------------------------------------------------------------------------------------'
