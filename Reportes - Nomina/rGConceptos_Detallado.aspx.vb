'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rGConceptos_Detallado"
'-------------------------------------------------------------------------------------------'
Partial Class rGConceptos_Detallado
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
			loComandoSeleccionar.AppendLine("SELECT		Grupos_Conceptos.Cod_Gru					AS Cod_Gru,	")
			loComandoSeleccionar.AppendLine("			Grupos_Conceptos.Nom_Gru					AS Nom_Gru,	")
			loComandoSeleccionar.AppendLine("			Grupos_Conceptos.Status						AS Status,	")
			loComandoSeleccionar.AppendLine("			Grupos_Conceptos.Comentario					AS Comentario, ")
			loComandoSeleccionar.AppendLine("			(CASE WHEN Grupos_Conceptos.Status = 'A'")
			loComandoSeleccionar.AppendLine("				THEN 'Activo' ")
			loComandoSeleccionar.AppendLine("				ELSE 'Inactivo' ")
			loComandoSeleccionar.AppendLine("			END)										AS Estatus,")
			loComandoSeleccionar.AppendLine("			Conceptos_Nomina.cod_con					AS cod_con,")
			loComandoSeleccionar.AppendLine("			Conceptos_Nomina.nom_con					AS nom_con")
			loComandoSeleccionar.AppendLine("FROM		Grupos_Conceptos")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Grupos_Conceptos")
			loComandoSeleccionar.AppendLine("		ON  Renglones_Grupos_Conceptos.cod_gru = Grupos_Conceptos.Cod_Gru")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Conceptos_Nomina")
			loComandoSeleccionar.AppendLine("		ON  Conceptos_Nomina.cod_con = Renglones_Grupos_Conceptos.cod_con")
			loComandoSeleccionar.AppendLine("WHERE	Grupos_Conceptos.Cod_Gru BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("		AND	" & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("		AND	Grupos_Conceptos.Status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			
			
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rGConceptos_Detallado", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrGConceptos_Detallado.ReportSource = loObjetoReporte

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
