'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rContratos_Detallados"
'-------------------------------------------------------------------------------------------'
Partial Class rContratos_Detallados
Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT		Contratos.Cod_Con					AS Cod_Con,	")
			loComandoSeleccionar.AppendLine("			Contratos.Nom_Con					AS Nom_Con,	")
			loComandoSeleccionar.AppendLine("			Contratos.Status					AS Status,	")
			loComandoSeleccionar.AppendLine("			Contratos.Comentario				AS Comentario, ")
			loComandoSeleccionar.AppendLine("			(CASE WHEN Contratos.Status = 'A'")
			loComandoSeleccionar.AppendLine("				THEN 'Activo' ")
			loComandoSeleccionar.AppendLine("				ELSE 'Inactivo' ")
			loComandoSeleccionar.AppendLine("			END)										AS Estatus,")
			loComandoSeleccionar.AppendLine("			Grupos_Conceptos.cod_gru					AS cod_gru,")
			loComandoSeleccionar.AppendLine("			Grupos_Conceptos.nom_gru					AS nom_gru")
			loComandoSeleccionar.AppendLine("FROM		Contratos")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Contratos")
			loComandoSeleccionar.AppendLine("		ON  Renglones_Contratos.Cod_Con = Contratos.Cod_Con")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Grupos_Conceptos")
			loComandoSeleccionar.AppendLine("		ON  Grupos_Conceptos.cod_gru = Renglones_Contratos.cod_gru")
			loComandoSeleccionar.AppendLine("WHERE	Contratos.Cod_Con BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("		AND	" & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("		AND	Contratos.Status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			
			
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rContratos_Detallados", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrContratos_Detallados.ReportSource = loObjetoReporte

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
