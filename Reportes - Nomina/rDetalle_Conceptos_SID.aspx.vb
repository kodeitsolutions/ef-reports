'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDetalle_Conceptos_SID"
'-------------------------------------------------------------------------------------------'
Partial Class rDetalle_Conceptos_SID
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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            'Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            Dim lcEstatus As String

            Select Case CStr(cusAplicacion.goReportes.paParametrosIniciales(3)).ToUpper().Trim()
                Case "A"
                    lcEstatus = "'Activos'"
                Case "I"
                    lcEstatus = "'Inactivos'"
                Case "S"
                    lcEstatus = "'Suspendidos'"
                Case "L"
                    lcEstatus = "'Liquidados'"
                Case Else
                    lcEstatus = "'Todos'"
            End Select



            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Trabajadores.Cod_Tra                                    AS Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                                    AS Empleado,")
            loConsulta.AppendLine("            SUM(CASE WHEN Conceptos_Nomina.tipo='Asignacion'")
            loConsulta.AppendLine("                        AND Renglones_Recibos.cod_con <> 'A011'")
            loConsulta.AppendLine("                        AND Renglones_Recibos.cod_con <> 'A013'")
            loConsulta.AppendLine("                THEN Renglones_Recibos.mon_net ")
            loConsulta.AppendLine("                ELSE 0 END)                                         AS Asignacion,")
            loConsulta.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loConsulta.AppendLine("                WHEN 'Deduccion' THEN Renglones_Recibos.mon_net")
            loConsulta.AppendLine("                WHEN 'Retencion' THEN Renglones_Recibos.mon_net")
            loConsulta.AppendLine("                ELSE 0 END)                                         AS Deduccion,")
            loConsulta.AppendLine("            SUM(CASE Conceptos_Nomina.tipo ")
            loConsulta.AppendLine("                WHEN 'Otro' THEN Renglones_Recibos.mon_net ")
            loConsulta.AppendLine("                ELSE 0 END)                                         AS Otro,")
            loConsulta.AppendLine("            SUM(CASE WHEN Renglones_Recibos.cod_con = 'A011'")
            loConsulta.AppendLine("                        OR Renglones_Recibos.cod_con = 'A013'")
            loConsulta.AppendLine("                THEN Renglones_Recibos.mon_net ")
            loConsulta.AppendLine("                ELSE 0 END)                                         AS Bono_Alimentacion,")
            loConsulta.AppendLine("                " & lcEstatus & "                                   AS Estatus")
            loConsulta.AppendLine("FROM        Renglones_Recibos")
            loConsulta.AppendLine("    JOIN    Recibos ")
            loConsulta.AppendLine("        ON  Recibos.Documento = Renglones_Recibos.Documento")
            loConsulta.AppendLine("        AND Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loConsulta.AppendLine("        AND Recibos.Cod_Con BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loConsulta.AppendLine("        AND Recibos.Status IN (" & lcParametro8Desde & ")")
            loConsulta.AppendLine("    JOIN    Conceptos_Nomina ")
            loConsulta.AppendLine("        ON  Conceptos_Nomina.Cod_Con = Renglones_Recibos.cod_con ")
            loConsulta.AppendLine("        AND Conceptos_Nomina.Cod_Con BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Conceptos_Nomina.Cod_Cla BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loConsulta.AppendLine("    RIGHT JOIN    Trabajadores ")
            loConsulta.AppendLine("        ON  Trabajadores.Cod_Tra = Recibos.cod_tra ")
            loConsulta.AppendLine("WHERE       Trabajadores.Tip_tra = 'Trabajador'")
            loConsulta.AppendLine("        AND Trabajadores.Cod_Tra BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Status IN (" & lcParametro3Desde & ")")
            loConsulta.AppendLine("        AND Trabajadores.Cod_Dep BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Suc BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loConsulta.AppendLine("GROUP BY    Trabajadores.Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra")
            loConsulta.AppendLine("ORDER BY	" & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDetalle_Conceptos_SID", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDetalle_Conceptos_SID.ReportSource = loObjetoReporte

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
' RJG: 13/06/13: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
