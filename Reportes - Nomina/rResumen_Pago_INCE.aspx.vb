'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rResumen_Pago_INCE"
'-------------------------------------------------------------------------------------------'
Partial Class rResumen_Pago_INCE
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
            loComandoSeleccionar.AppendLine("			CASE ")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(1,2,3) THEN 1")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(4,5,6) THEN 2")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(7,8,9) THEN 3")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(10,11,12) THEN 4")
            loComandoSeleccionar.AppendLine("			END														 AS Trimestre,")
            loComandoSeleccionar.AppendLine("            COUNT(DISTINCT Recibos.Cod_Tra)                          AS Trabajadores,")
            loComandoSeleccionar.AppendLine("			SUM(CASE Conceptos_Nomina.tipo")
            loComandoSeleccionar.AppendLine("					WHEN 'Asignacion' THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("					else -Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("				END)												AS Monto_Devengado,")
            loComandoSeleccionar.AppendLine("			(SELECT TOP 1 val_num")
            loComandoSeleccionar.AppendLine("           FROM Constantes_Locales")
            loComandoSeleccionar.AppendLine("			WHERE COD_CON ='U004')*")
            loComandoSeleccionar.AppendLine("			SUM(CASE Conceptos_Nomina.tipo")
            loComandoSeleccionar.AppendLine("					WHEN 'Asignacion' THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("					else -Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("				END)/100											AS Aporte")
            loComandoSeleccionar.AppendLine("FROM        Renglones_Recibos")
            loComandoSeleccionar.AppendLine("    JOIN    Recibos")
            loComandoSeleccionar.AppendLine("        ON  Recibos.Documento = Renglones_Recibos.Documento")
            loComandoSeleccionar.AppendLine("    JOIN    Conceptos_Nomina")
            loComandoSeleccionar.AppendLine("        ON  Conceptos_Nomina.cod_con = Renglones_Recibos.cod_con ")
            loComandoSeleccionar.AppendLine("    JOIN    Trabajadores")
            loComandoSeleccionar.AppendLine("        ON  Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loComandoSeleccionar.AppendLine("WHERE       Conceptos_Nomina.tipo <> 'otro'")
            loComandoSeleccionar.AppendLine("    AND     Conceptos_Nomina.acumulados = 1")
            loComandoSeleccionar.AppendLine("    AND     Recibos.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("    AND     Recibos.Fecha BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("    AND     Recibos.Documento BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("    AND     Recibos.Cod_Con BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("    AND     Trabajadores.Cod_Tra BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("    AND     Trabajadores.Cod_Con BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("    AND     Recibos.Cod_Suc BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("    AND     Recibos.Cod_Rev BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY    RIGHT('0000'+CAST(YEAR(Recibos.Fecha) as VARCHAR(4)),4),")
            loComandoSeleccionar.AppendLine("			CASE ")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(1,2,3) THEN 1")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(4,5,6) THEN 2")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(7,8,9) THEN 3")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(10,11,12) THEN 4")
            loComandoSeleccionar.AppendLine("            End")
            loComandoSeleccionar.AppendLine("ORDER BY    RIGHT('0000'+CAST(YEAR(Recibos.Fecha) as VARCHAR(4)),4),")
            loComandoSeleccionar.AppendLine("			CASE ")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(1,2,3) THEN 1")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(4,5,6) THEN 2")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(7,8,9) THEN 3")
            loComandoSeleccionar.AppendLine("				WHEN MONTH(Recibos.Fecha) IN(10,11,12) THEN 4")
            loComandoSeleccionar.AppendLine("            End")


            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rResumen_Pago_INCE", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrResumen_Pago_INCE.ReportSource = loObjetoReporte

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
' EAG: 02/10/15: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
