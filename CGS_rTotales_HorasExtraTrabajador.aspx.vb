'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rTotales_HorasExtraTrabajador"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rTotales_HorasExtraTrabajador
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE @lcCodConc_Desde AS VARCHAR(4) = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @lcCodConc_Hasta AS VARCHAR(4) = " & lcParametro0Hasta)
            loConsulta.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro1Desde)
            loConsulta.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro1Hasta)
            loConsulta.AppendLine("DECLARE @lcCodCont_Desde AS VARCHAR(2) = " & lcParametro2Desde)
            loConsulta.AppendLine("DECLARE @lcCodCont_Hasta AS VARCHAR(2) = " & lcParametro2Hasta)
            loConsulta.AppendLine("DECLARE @lcCodTra_Desde AS VARCHAR(9) =  " & lcParametro3Desde)
            loConsulta.AppendLine("DECLARE @lcCodTra_Hasta AS VARCHAR(9) =  " & lcParametro3Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Conceptos_Nomina.Cod_Con        AS Cod_Con,")
            loConsulta.AppendLine("        Conceptos_Nomina.Nom_Con         AS Nom_Con,")
            loConsulta.AppendLine("        Trabajadores.Cod_Tra             AS Cod_Tra,")
            loConsulta.AppendLine("        Trabajadores.Nom_Tra             AS Nom_Tra,")
            loConsulta.AppendLine("        SUM(Renglones_Recibos.Mon_Net)   AS Mon_Net,")
            loConsulta.AppendLine("        SUM(Renglones_Recibos.Val_Num)   AS Val_Num")
            loConsulta.AppendLine("FROM Renglones_Recibos")
            loConsulta.AppendLine("    JOIN Recibos ON Recibos.Documento = Renglones_Recibos.Documento")
            loConsulta.AppendLine("    JOIN Conceptos_Nomina ON Conceptos_Nomina.Cod_Con = Renglones_Recibos.cod_con ")
            loConsulta.AppendLine("        AND Conceptos_Nomina.Cod_Con IN ('B001', 'B002', 'B103', 'B601','B602','B605','B015','B016','A038','A039') ")
            loConsulta.AppendLine("    JOIN Trabajadores ON Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loConsulta.AppendLine("WHERE Conceptos_Nomina.Cod_Con BETWEEN @lcCodConc_Desde AND @lcCodConc_Hasta")
            loConsulta.AppendLine("	AND Recibos.Fecha BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loConsulta.AppendLine("    AND Recibos.Cod_Con BETWEEN @lcCodCont_Desde AND @lcCodCont_Hasta")
            loConsulta.AppendLine("    AND Trabajadores.Cod_Tra BETWEEN @lcCodTra_Desde AND @lcCodTra_Hasta")
            loConsulta.AppendLine("    AND Recibos.Status IN ('Confirmado', 'Procesado')")
            loConsulta.AppendLine("GROUP BY Conceptos_Nomina.Cod_Con,Conceptos_Nomina.Nom_Con, Conceptos_Nomina.Tipo,Trabajadores.Cod_Tra,Trabajadores.Nom_Tra")
            loConsulta.AppendLine("ORDER BY Trabajadores.Cod_Tra, Cod_Con ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rTotales_HorasExtraTrabajador", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rTotales_HorasExtraTrabajador.ReportSource = loObjetoReporte

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
' RJG: 22/04/15: Codigo inicial.							                                '
'-------------------------------------------------------------------------------------------'
