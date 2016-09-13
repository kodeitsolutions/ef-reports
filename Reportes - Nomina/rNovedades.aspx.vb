'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rNovedades"
'-------------------------------------------------------------------------------------------'
Partial Class rNovedades
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
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim loConsulta As New StringBuilder()

			loConsulta.AppendLine("")
			loConsulta.AppendLine("SELECT      Novedades.documento             AS documento, ")
			loConsulta.AppendLine("            Novedades.fecha                 AS fecha,")
			loConsulta.AppendLine("            Novedades.status                AS estatus, ")
			loConsulta.AppendLine("            Novedades.cod_tra               AS cod_tra, ")
			loConsulta.AppendLine("            Trabajadores.nom_tra            AS nom_tra,")
			loConsulta.AppendLine("            Novedades.fec_ini               AS fec_ini,")
			loConsulta.AppendLine("            Novedades.fec_fin               AS fec_fin,")
			loConsulta.AppendLine("            Novedades.horas                 AS horas,")
			loConsulta.AppendLine("            Novedades.remunerado            AS remunerado,")
			loConsulta.AppendLine("            Novedades.justificado           AS justificado,")
			loConsulta.AppendLine("            Novedades.aut_por               AS aut_por,")
			loConsulta.AppendLine("            Novedades.motivo                AS Motivo,")
			loConsulta.AppendLine("            Novedades.concepto              AS Concepto,")
			loConsulta.AppendLine("            Novedades.cod_rev               AS Cod_Rev")
			loConsulta.AppendLine("FROM        Novedades")
			loConsulta.AppendLine("    JOIN    Trabajadores ")
			loConsulta.AppendLine("        ON  Trabajadores.cod_tra = Novedades.cod_tra")
			loConsulta.AppendLine("    JOIN    Trabajadores AS Autorizacion ")
			loConsulta.AppendLine("        ON  Autorizacion.cod_tra = Novedades.cod_tra")
			loConsulta.AppendLine("    JOIN    Conceptos_Nomina ")
			loConsulta.AppendLine("        ON  Conceptos_Nomina.cod_con = Novedades.concepto")
			loConsulta.AppendLine("    JOIN    Motivos ")
			loConsulta.AppendLine("        ON  Motivos.cod_mot = Novedades.motivo")
			loConsulta.AppendLine("WHERE		Novedades.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			loConsulta.AppendLine("		AND	Novedades.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
			loConsulta.AppendLine("		AND	Novedades.Cod_Tra BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loConsulta.AppendLine("		AND	Novedades.Aut_Por BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loConsulta.AppendLine("		AND	Novedades.Status IN (" & lcParametro3Desde & ")")
			loConsulta.AppendLine("		AND	Novedades.Motivo BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
			loConsulta.AppendLine("		AND	Novedades.Concepto BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
			loConsulta.AppendLine("		AND	Novedades.Cod_Rev BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
			loConsulta.AppendLine("ORDER BY	" & lcOrdenamiento)
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")

			Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loConsulta.ToString())
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
			
			
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rNovedades", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrNovedades.ReportSource = loObjetoReporte

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
' RJG: 16/02/13: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
