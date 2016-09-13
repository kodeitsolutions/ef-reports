'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTotales_HorasExtraTrabajador"
'-------------------------------------------------------------------------------------------'
Partial Class rTotales_HorasExtraTrabajador
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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Conceptos_Nomina.Cod_Con                AS Cod_Con,")
            loConsulta.AppendLine("            Conceptos_Nomina.Nom_Con                AS Nom_Con,")
            loConsulta.AppendLine("            Conceptos_Nomina.Tipo                   AS Tipo,")
            loConsulta.AppendLine("            Trabajadores.Cod_Tra                    AS Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                    AS Nom_Tra,")
            loConsulta.AppendLine("            SUM(Renglones_Recibos.Mon_Net)          AS Mon_Net,")
            loConsulta.AppendLine("            SUM(Renglones_Recibos.Val_Num)          AS Val_Num")
            loConsulta.AppendLine("FROM        Renglones_Recibos")
            loConsulta.AppendLine("    JOIN    Recibos ")
            loConsulta.AppendLine("        ON  Recibos.Documento = Renglones_Recibos.Documento")
            loConsulta.AppendLine("    JOIN    Conceptos_Nomina ")
            loConsulta.AppendLine("        ON  Conceptos_Nomina.Cod_Con = Renglones_Recibos.cod_con ")
            loConsulta.AppendLine("        AND Conceptos_Nomina.Cod_Con IN ('B001', 'B002', 'B103', 'B601','B602','B605','B015','B016') ")
            loConsulta.AppendLine("    JOIN    Trabajadores ")
            loConsulta.AppendLine("        ON  Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loConsulta.AppendLine("WHERE		 Conceptos_Nomina.Cod_Con BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loConsulta.AppendLine("        AND Recibos.Cod_Con BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Dep BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Car BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Tra BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loConsulta.AppendLine("        AND Conceptos_Nomina.Cod_Cla BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Suc BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loConsulta.AppendLine("        AND Recibos.Cod_Mon BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta)
            loConsulta.AppendLine("        AND Recibos.Status IN ('Confirmado', 'Procesado')")
            loConsulta.AppendLine("GROUP BY    conceptos_Nomina.Cod_Con,")
            loConsulta.AppendLine("            conceptos_Nomina.Nom_Con,")
            loConsulta.AppendLine("            conceptos_Nomina.Tipo,")
            loConsulta.AppendLine("            Trabajadores.Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra")
            loConsulta.AppendLine("ORDER BY    Trabajadores.Cod_Tra, Cod_Con ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTotales_HorasExtraTrabajador", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTotales_HorasExtraTrabajador.ReportSource = loObjetoReporte

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
