'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRecibos_Trabajador"
'-------------------------------------------------------------------------------------------'
Partial Class rRecibos_Trabajador
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
			Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro54Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("SELECT		Recibos.Documento		AS Documento,	")
			loComandoSeleccionar.AppendLine("			Recibos.fecha			AS Fecha,		")
			loComandoSeleccionar.AppendLine("			Recibos.[Status]		AS Estatus,		")
			loComandoSeleccionar.AppendLine("			Recibos.Cod_Con			AS Cod_Con,		")
			loComandoSeleccionar.AppendLine("			Recibos.Cod_Tra			AS Cod_Tra,		")
			loComandoSeleccionar.AppendLine("			Trabajadores.Nom_Tra	AS Nom_Tra,		")
			loComandoSeleccionar.AppendLine("			Recibos.Fec_Ini			AS Fec_Ini,		")
			loComandoSeleccionar.AppendLine("			Recibos.Fec_Fin			AS Fec_Fin,		")
			loComandoSeleccionar.AppendLine("			Recibos.Cod_Rev			AS Cod_Rev,		")
			loComandoSeleccionar.AppendLine("			Recibos.Comentario		AS Comentario,	")
			loComandoSeleccionar.AppendLine("			Recibos.Mon_Asi			AS Mon_Asi,		")
			loComandoSeleccionar.AppendLine("			Recibos.Mon_Ded			AS Mon_Ded,		")
			loComandoSeleccionar.AppendLine("			Recibos.Mon_Ret			AS Mon_Ret,		")
			loComandoSeleccionar.AppendLine("			Recibos.Mon_Otr1		AS Mon_Otr1,	")
			loComandoSeleccionar.AppendLine("			Recibos.Mon_Net			AS Mon_Net		")
			loComandoSeleccionar.AppendLine("FROM		Recibos ")
			loComandoSeleccionar.AppendLine("	JOIN	Trabajadores ")
			loComandoSeleccionar.AppendLine("		ON	Trabajadores.Cod_Tra = Recibos.Cod_Tra")
			loComandoSeleccionar.AppendLine("WHERE		Recibos.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("		AND	Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Tra BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Con BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("		AND	Recibos.Status IN (" & lcParametro4Desde & ")")
			loComandoSeleccionar.AppendLine("		AND	Recibos.Cod_Rev BETWEEN " & lcParametro54Desde & " AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)

			Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			
			
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRecibos_Trabajador", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrRecibos_Trabajador.ReportSource = loObjetoReporte

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
