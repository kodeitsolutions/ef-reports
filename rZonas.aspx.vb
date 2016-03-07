'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rZonas"
'-------------------------------------------------------------------------------------------'
Partial Class rZonas
Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

			Try
				Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
				Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
				Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

				Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
				Dim loComandoSeleccionar As New StringBuilder()

				loComandoSeleccionar.AppendLine("SELECT		Cod_Zon, ")
				loComandoSeleccionar.AppendLine("			Nom_Zon, ")
				loComandoSeleccionar.AppendLine("			Status, ")
				loComandoSeleccionar.AppendLine("			(CASE WHEN Status = 'A' THEN 'Activo' ELSE 'Inactivo' END) AS Status_Zonas ")
				loComandoSeleccionar.AppendLine("FROM		Zonas ")
				loComandoSeleccionar.AppendLine("WHERE	Cod_Zon BETWEEN " & lcParametro0Desde)
				loComandoSeleccionar.AppendLine("		AND	" & lcParametro0Hasta)
				loComandoSeleccionar.AppendLine("		AND	Status IN (" & lcParametro1Desde & ")")
				loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)

						Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rZonas", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrZonas.ReportSource = loObjetoReporte

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
' MJP: 07/07/08: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' MJP: 14/07/08: Creación objeto que cierra el archivo de reporte. Se agregó el filtro de	'
'				 Estatus.																	'
'-------------------------------------------------------------------------------------------'
' MVP: 04/08/08: Cambios para multi idioma, mensaje de error y clase padre.					'
'-------------------------------------------------------------------------------------------'
' RJG: 16/02/13: Ajuste visual. Se agregaron comentarios y estándar faltantes.				'
'-------------------------------------------------------------------------------------------'
