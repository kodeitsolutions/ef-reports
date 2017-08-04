'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rTotales_HorasExtra"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rTotales_HorasExtra
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE @lcCodCon_Desde AS VARCHAR(4) = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @lcCodCon_Hasta AS VARCHAR(4) = " & lcParametro0Hasta)
            loConsulta.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro1Desde)
            loConsulta.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro1Hasta)
            loConsulta.AppendLine("DECLARE @lcCodCnt_Desde AS VARCHAR(2) = " & lcParametro3Desde)
            loConsulta.AppendLine("DECLARE @lcCodCnt_Hasta AS VARCHAR(2) = " & lcParametro3Hasta)
            loConsulta.AppendLine("DECLARE @lcCodDep_Desde AS VARCHAR(10) = " & lcParametro4Desde)
            loConsulta.AppendLine("DECLARE @lcCodDep_Hasta AS VARCHAR(10) = " & lcParametro4Hasta)
            loConsulta.AppendLine("DECLARE @lcCodTra_Desde AS VARCHAR(10) = " & lcParametro5Desde)
            loConsulta.AppendLine("DECLARE @lcCodTra_Hasta AS VARCHAR(10) = " & lcParametro5Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   CASE WHEN Conceptos_Nomina.Cod_Con IN ('B001', 'B015','A038')")
            loConsulta.AppendLine("              THEN 'Total Horas Extra Diurnas'")
            loConsulta.AppendLine("              ELSE 'Total Horas Extra Nocturnas'")
            loConsulta.AppendLine("          END                                      AS Concepto,")
            loConsulta.AppendLine("          SUM(Renglones_Recibos.Val_Num)           AS Horas,")
            loConsulta.AppendLine("          SUM(Renglones_Recibos.Mon_Net)           AS Monto")
            loConsulta.AppendLine("FROM Renglones_Recibos")
            loConsulta.AppendLine("	    JOIN Recibos ON Recibos.Documento = Renglones_Recibos.Documento")
            loConsulta.AppendLine("     JOIN Conceptos_Nomina ON Conceptos_Nomina.Cod_Con = Renglones_Recibos.Cod_Con ")
            loConsulta.AppendLine("     JOIN Trabajadores ON Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loConsulta.AppendLine("WHERE Conceptos_Nomina.Cod_Con IN ('B001', 'B002', 'B015', 'B016','A038','A039')")
            loConsulta.AppendLine("     AND Conceptos_Nomina.Cod_Con BETWEEN @lcCodCon_Desde AND @lcCodCon_Hasta")
            loConsulta.AppendLine("     AND Recibos.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loConsulta.AppendLine("     AND Conceptos_Nomina.Tipo IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine("     AND Recibos.Cod_Con BETWEEN @lcCodCnt_Desde And @lcCodCnt_Hasta")
            loConsulta.AppendLine("     AND Trabajadores.Cod_Dep BETWEEN @lcCodDep_Desde And @lcCodDep_Hasta")
            loConsulta.AppendLine("     AND Trabajadores.Cod_Tra BETWEEN @lcCodTra_Desde And @lcCodTra_Hasta")
            loConsulta.AppendLine("     AND Recibos.Status In ('Confirmado', 'Procesado')")
            loConsulta.AppendLine("GROUP BY CASE WHEN Conceptos_Nomina.Cod_Con IN ('B001', 'B015','A038','A039')")
            loConsulta.AppendLine("              THEN 'Total Horas Extra Diurnas'")
            loConsulta.AppendLine("              ELSE 'Total Horas Extra Nocturnas'")
            loConsulta.AppendLine("         END, Conceptos_Nomina.Cod_Con")
            loConsulta.AppendLine("ORDER BY Concepto ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rTotales_HorasExtra", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rTotales_HorasExtra.ReportSource = loObjetoReporte

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
' RJG: 22/06/14: Codigo inicial, a partir de rTotales_Conceptos.							'
'-------------------------------------------------------------------------------------------'
