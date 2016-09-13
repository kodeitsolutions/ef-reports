'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTotales_Turnos"
'-------------------------------------------------------------------------------------------'
Partial Class rTotales_Turnos
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
            Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT      Turnos_Nomina.nom_tur,")
            loComandoSeleccionar.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loComandoSeleccionar.AppendLine("                WHEN 'Asignacion'")
            loComandoSeleccionar.AppendLine("                THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("                ELSE 0")
            loComandoSeleccionar.AppendLine("            END)                                   AS Asignacion,")
            loComandoSeleccionar.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loComandoSeleccionar.AppendLine("                WHEN 'Deduccion'")
            loComandoSeleccionar.AppendLine("                THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("                ELSE 0")
            loComandoSeleccionar.AppendLine("            END)                                   AS Deduccion,")
            loComandoSeleccionar.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loComandoSeleccionar.AppendLine("                WHEN 'Retencion'")
            loComandoSeleccionar.AppendLine("                THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("                ELSE 0")
            loComandoSeleccionar.AppendLine("            END)                                   AS Retencion,")

            loComandoSeleccionar.AppendLine("            -SUM(CASE Conceptos_Nomina.tipo ")
            loComandoSeleccionar.AppendLine("                WHEN 'Deduccion' ")
            loComandoSeleccionar.AppendLine("                THEN Renglones_Recibos.mon_net ")
            loComandoSeleccionar.AppendLine("                ELSE 0 ")
            loComandoSeleccionar.AppendLine("            END) - ")
            loComandoSeleccionar.AppendLine("            SUM(CASE Conceptos_Nomina.tipo ")
            loComandoSeleccionar.AppendLine("                WHEN 'Retencion' ")
            loComandoSeleccionar.AppendLine("                THEN Renglones_Recibos.mon_net ")
            loComandoSeleccionar.AppendLine("                ELSE 0 ")
            loComandoSeleccionar.AppendLine("            END)                                   AS Deduc_ret,  ")
            loComandoSeleccionar.AppendLine("            (CASE WHEN SUM(CASE Conceptos_Nomina.tipo  ")
            loComandoSeleccionar.AppendLine("							WHEN 'Otro'  ")
            loComandoSeleccionar.AppendLine("							THEN Renglones_Recibos.mon_net  ")
            loComandoSeleccionar.AppendLine("							ELSE 0  ")
            loComandoSeleccionar.AppendLine("						END)   <= 0      ")
            loComandoSeleccionar.AppendLine("					THEN     SUM(CASE Conceptos_Nomina.tipo  ")
            loComandoSeleccionar.AppendLine("									WHEN 'Asignacion'  ")
            loComandoSeleccionar.AppendLine("									THEN Renglones_Recibos.mon_net  ")
            loComandoSeleccionar.AppendLine("									ELSE 0  ")
            loComandoSeleccionar.AppendLine("							 END)   -    ")
            loComandoSeleccionar.AppendLine("							SUM(CASE Conceptos_Nomina.tipo  ")
            loComandoSeleccionar.AppendLine("								WHEN 'Deduccion'  ")
            loComandoSeleccionar.AppendLine("								THEN Renglones_Recibos.mon_net  ")
            loComandoSeleccionar.AppendLine("								ELSE 0  ")
            loComandoSeleccionar.AppendLine("							END)   -                                 ")
            loComandoSeleccionar.AppendLine("							SUM(CASE Conceptos_Nomina.tipo  ")
            loComandoSeleccionar.AppendLine("								WHEN 'Retencion'  ")
            loComandoSeleccionar.AppendLine("								THEN Renglones_Recibos.mon_net  ")
            loComandoSeleccionar.AppendLine("								ELSE 0  ")
            loComandoSeleccionar.AppendLine("							END)       ")
            loComandoSeleccionar.AppendLine("					ELSE    SUM(CASE Conceptos_Nomina.tipo  ")
            loComandoSeleccionar.AppendLine("									WHEN 'Otro'  ")
            loComandoSeleccionar.AppendLine("									THEN Renglones_Recibos.mon_net  ")
            loComandoSeleccionar.AppendLine("									ELSE 0  ")
            loComandoSeleccionar.AppendLine("								END)                    ")
            loComandoSeleccionar.AppendLine("			END)*100/  ")
            loComandoSeleccionar.AppendLine("			SUM((CASE WHEN SUM(CASE Conceptos_Nomina.tipo  ")
            loComandoSeleccionar.AppendLine("									WHEN 'Otro'  ")
            loComandoSeleccionar.AppendLine("									THEN Renglones_Recibos.mon_net  ")
            loComandoSeleccionar.AppendLine("									ELSE 0  ")
            loComandoSeleccionar.AppendLine("								END)   <= 0      ")
            loComandoSeleccionar.AppendLine("						THEN     SUM(CASE Conceptos_Nomina.tipo  ")
            loComandoSeleccionar.AppendLine("											WHEN 'Asignacion'  ")
            loComandoSeleccionar.AppendLine("											THEN Renglones_Recibos.mon_net  ")
            loComandoSeleccionar.AppendLine("											ELSE 0  ")
            loComandoSeleccionar.AppendLine("									 END)   -    ")
            loComandoSeleccionar.AppendLine("							SUM(CASE Conceptos_Nomina.tipo  ")
            loComandoSeleccionar.AppendLine("									WHEN 'Deduccion'  ")
            loComandoSeleccionar.AppendLine("									THEN Renglones_Recibos.mon_net  ")
            loComandoSeleccionar.AppendLine("									ELSE 0  ")
            loComandoSeleccionar.AppendLine("								END)   -                                 ")
            loComandoSeleccionar.AppendLine("							SUM(CASE Conceptos_Nomina.tipo  ")
            loComandoSeleccionar.AppendLine("									WHEN 'Retencion'  ")
            loComandoSeleccionar.AppendLine("									THEN Renglones_Recibos.mon_net  ")
            loComandoSeleccionar.AppendLine("									ELSE 0  ")
            loComandoSeleccionar.AppendLine("								END)       ")
            loComandoSeleccionar.AppendLine("					ELSE    SUM(CASE Conceptos_Nomina.tipo  ")
            loComandoSeleccionar.AppendLine("									WHEN 'Otro'  ")
            loComandoSeleccionar.AppendLine("									THEN Renglones_Recibos.mon_net  ")
            loComandoSeleccionar.AppendLine("									ELSE 0  ")
            loComandoSeleccionar.AppendLine("								END)                    ")
            loComandoSeleccionar.AppendLine("			END)) OVER() por_neto,  ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN SUM(CASE Conceptos_Nomina.tipo ")
            loComandoSeleccionar.AppendLine("							WHEN 'Otro' ")
            loComandoSeleccionar.AppendLine("							THEN Renglones_Recibos.mon_net ")
            loComandoSeleccionar.AppendLine("							ELSE 0 ")
            loComandoSeleccionar.AppendLine("						END)   <= 0     ")
            loComandoSeleccionar.AppendLine("					THEN     SUM(CASE Conceptos_Nomina.tipo ")
            loComandoSeleccionar.AppendLine("									WHEN 'Asignacion' ")
            loComandoSeleccionar.AppendLine("									THEN Renglones_Recibos.mon_net ")
            loComandoSeleccionar.AppendLine("									ELSE 0 ")
            loComandoSeleccionar.AppendLine("							 END)   -   ")
            loComandoSeleccionar.AppendLine("							SUM(CASE Conceptos_Nomina.tipo ")
            loComandoSeleccionar.AppendLine("								WHEN 'Deduccion' ")
            loComandoSeleccionar.AppendLine("								THEN Renglones_Recibos.mon_net ")
            loComandoSeleccionar.AppendLine("								ELSE 0 ")
            loComandoSeleccionar.AppendLine("							END)   -                                ")
            loComandoSeleccionar.AppendLine("							SUM(CASE Conceptos_Nomina.tipo ")
            loComandoSeleccionar.AppendLine("								WHEN 'Retencion' ")
            loComandoSeleccionar.AppendLine("								THEN Renglones_Recibos.mon_net ")
            loComandoSeleccionar.AppendLine("								ELSE 0 ")
            loComandoSeleccionar.AppendLine("							END)      ")
            loComandoSeleccionar.AppendLine("					ELSE    SUM(CASE Conceptos_Nomina.tipo ")
            loComandoSeleccionar.AppendLine("									WHEN 'Otro' ")
            loComandoSeleccionar.AppendLine("									THEN Renglones_Recibos.mon_net ")
            loComandoSeleccionar.AppendLine("									ELSE 0 ")
            loComandoSeleccionar.AppendLine("								END)       ")
            loComandoSeleccionar.AppendLine("			END) AS Total_neto, ")
            loComandoSeleccionar.AppendLine("            COUNT(DISTINCT Trabajadores.cod_tra)   AS  num_tra,")
            loComandoSeleccionar.AppendLine("			CAST(COUNT(DISTINCT Trabajadores.cod_tra)AS DECIMAL(28,10))*100/SUM(COUNT(DISTINCT Trabajadores.cod_tra)) OVER() por_tra ")
            loComandoSeleccionar.AppendLine("FROM        Renglones_Recibos")
            loComandoSeleccionar.AppendLine("    JOIN    Recibos ")
            loComandoSeleccionar.AppendLine("        ON  Recibos.Documento = Renglones_Recibos.Documento")
            loComandoSeleccionar.AppendLine("    JOIN    Conceptos_Nomina ")
            loComandoSeleccionar.AppendLine("        ON  Conceptos_Nomina.cod_con = Renglones_Recibos.cod_con ")
            'loComandoSeleccionar.AppendLine("        AND Conceptos_Nomina.Tipo <> 'Otro' ")
            loComandoSeleccionar.AppendLine("    JOIN    Trabajadores ")
            loComandoSeleccionar.AppendLine("        ON  Trabajadores.cod_tra = Recibos.cod_tra ")
            loComandoSeleccionar.AppendLine("    JOIN    Departamentos_Nomina ")
            loComandoSeleccionar.AppendLine("        ON  Departamentos_Nomina.Cod_Dep = Trabajadores.Cod_Dep ")
            loComandoSeleccionar.AppendLine("    JOIN    Turnos_Nomina ")
            loComandoSeleccionar.AppendLine("        ON  Turnos_Nomina.Cod_tur = Trabajadores.Cod_tur ")
            loComandoSeleccionar.AppendLine("WHERE		 Conceptos_Nomina.Cod_Con BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("        AND Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("        AND Conceptos_Nomina.Tipo IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("        AND Recibos.Cod_Con BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("        AND Departamentos_Nomina.Cod_Dep BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("        AND Trabajadores.Cod_Suc BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("        AND Trabajadores.Cod_Tra BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("        AND Conceptos_Nomina.Tipo BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("        AND Recibos.Status IN (" & lcParametro8Desde & ")")
            loComandoSeleccionar.AppendLine("GROUP BY    Turnos_Nomina.nom_tur")
            loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTotales_Turnos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTotales_Turnos.ReportSource = loObjetoReporte

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
' EAG: 07/09/15: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
