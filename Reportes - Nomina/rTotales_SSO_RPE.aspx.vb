'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTotales_SSO_RPE"
'-------------------------------------------------------------------------------------------'
Partial Class rTotales_SSO_RPE
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))    'Recibo
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))    'Contrato del Recibo
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))    'Trabajador
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))    'Contrato del Trabajador
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))    'Sucursal
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))    'Revisión
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT      RIGHT('0000'+CAST(YEAR(Recibos.Fecha) AS VARCHAR(4)),4)  AS Año,")
            loComandoSeleccionar.AppendLine("            RIGHT('00'+CAST(MONTH(Recibos.Fecha) AS VARCHAR(4)),2)   AS Mes,")
            loComandoSeleccionar.AppendLine("            COUNT(DISTINCT Recibos.Cod_Tra)                          AS Trabajadores,")
            loComandoSeleccionar.AppendLine("            SUM(CASE Conceptos_Nomina.Cod_Con")
            loComandoSeleccionar.AppendLine("                WHEN 'R001'")
            loComandoSeleccionar.AppendLine("                THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("                ELSE 0")
            loComandoSeleccionar.AppendLine("            END)                                    AS Retencion_SSO,")
            loComandoSeleccionar.AppendLine("            SUM(CASE Conceptos_Nomina.Cod_Con")
            loComandoSeleccionar.AppendLine("                WHEN 'R002'")
            loComandoSeleccionar.AppendLine("                THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("                ELSE 0")
            loComandoSeleccionar.AppendLine("            END)                                    AS Retencion_RPE,")
            loComandoSeleccionar.AppendLine("            SUM(CASE Conceptos_Nomina.Cod_Con")
            loComandoSeleccionar.AppendLine("                WHEN 'U001'")
            loComandoSeleccionar.AppendLine("                THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("                ELSE 0")
            loComandoSeleccionar.AppendLine("            END)                                    AS Aporte_SSO,")
            loComandoSeleccionar.AppendLine("            SUM(CASE Conceptos_Nomina.Cod_Con")
            loComandoSeleccionar.AppendLine("                WHEN 'U002'")
            loComandoSeleccionar.AppendLine("                THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("                ELSE 0")
            loComandoSeleccionar.AppendLine("            END)                                    AS Aporte_RPE,")
            loComandoSeleccionar.AppendLine("            SUM(Renglones_Recibos.mon_net)          AS Total")
            loComandoSeleccionar.AppendLine("FROM        Renglones_Recibos")
            loComandoSeleccionar.AppendLine("    JOIN    Recibos")
            loComandoSeleccionar.AppendLine("        ON  Recibos.Documento = Renglones_Recibos.Documento")
            loComandoSeleccionar.AppendLine("    JOIN    Conceptos_Nomina")
            loComandoSeleccionar.AppendLine("        ON  Conceptos_Nomina.cod_con = Renglones_Recibos.cod_con ")
            loComandoSeleccionar.AppendLine("    JOIN    Trabajadores")
            loComandoSeleccionar.AppendLine("        ON  Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loComandoSeleccionar.AppendLine("WHERE       Conceptos_Nomina.Cod_Con IN ('R001', 'R002', 'U001', 'U002')")
            loComandoSeleccionar.AppendLine("    AND     Recibos.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("    AND     Recibos.Fecha BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("    AND     Recibos.Documento BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("    AND     Recibos.Cod_Con BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("    AND     Trabajadores.Cod_Tra BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("    AND     Trabajadores.Cod_Con BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("    AND     Recibos.Cod_Suc BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("    AND     Recibos.Cod_Rev BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY    YEAR(Recibos.Fecha),")
            loComandoSeleccionar.AppendLine("            MONTH(Recibos.Fecha)")
            loComandoSeleccionar.AppendLine("ORDER BY    YEAR(Recibos.Fecha) ASC,")
            loComandoSeleccionar.AppendLine("            MONTH(Recibos.Fecha) ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTotales_SSO_RPE", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTotales_SSO_RPE.ReportSource = loObjetoReporte

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
' RJG: 04/09/15: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
