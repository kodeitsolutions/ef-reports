'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rVacaciones_Trabajador"
'-------------------------------------------------------------------------------------------'
Partial Class rVacaciones_Trabajador
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
			
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("SELECT		Vacaciones.Documento	AS Documento,	")
			loComandoSeleccionar.AppendLine("			Vacaciones.fecha		AS Fecha,		")
			loComandoSeleccionar.AppendLine("			Vacaciones.[Status]		AS Estatus,		")
			loComandoSeleccionar.AppendLine("			Vacaciones.Cod_Tra		AS Cod_Tra,		")
			loComandoSeleccionar.AppendLine("			Trabajadores.Nom_Tra	AS Nom_Tra,		")
			loComandoSeleccionar.AppendLine("			Vacaciones.Fec_Ini		AS Fec_Ini,		")
			loComandoSeleccionar.AppendLine("			Vacaciones.Fec_Fin		AS Fec_Fin,		")
			loComandoSeleccionar.AppendLine("			Vacaciones.Fec_Rei		AS Fec_Rei,		")
			loComandoSeleccionar.AppendLine("			Vacaciones.Dias			AS Dias,		")
			loComandoSeleccionar.AppendLine("			Vacaciones.Cod_Rev		AS Cod_Rev,		")
			loComandoSeleccionar.AppendLine("			Vacaciones.Motivo		AS Motivo,		")
			loComandoSeleccionar.AppendLine("			Vacaciones.Comentario	AS Comentario	")
			loComandoSeleccionar.AppendLine("FROM		Vacaciones ")
			loComandoSeleccionar.AppendLine("	JOIN	Trabajadores ")
			loComandoSeleccionar.AppendLine("		ON	Trabajadores.cod_Tra = Vacaciones.Cod_Tra")
			loComandoSeleccionar.AppendLine("WHERE		Vacaciones.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("		AND	Vacaciones.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("		AND	Vacaciones.Cod_Tra BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("		AND	Vacaciones.Status IN (" & lcParametro3Desde & ")")
			loComandoSeleccionar.AppendLine("		AND	Vacaciones.Cod_Rev BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)

			Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			
			
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rVacaciones_Trabajador", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrVacaciones_Trabajador.ReportSource = loObjetoReporte

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
