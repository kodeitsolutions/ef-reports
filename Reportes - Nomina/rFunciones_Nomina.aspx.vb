'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rFunciones_Nomina"
'-------------------------------------------------------------------------------------------'
Partial Class rFunciones_Nomina
Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("SELECT		Cod_Fun,	")
			loComandoSeleccionar.AppendLine("			Nom_Fun,	")
			loComandoSeleccionar.AppendLine("			Status,		")
			loComandoSeleccionar.AppendLine("			Comentario, ")
			loComandoSeleccionar.AppendLine("			Val_Fun,	")
			loComandoSeleccionar.AppendLine("			(CASE WHEN Status = 'A'")
			loComandoSeleccionar.AppendLine("				THEN 'Activo' ")
			loComandoSeleccionar.AppendLine("				ELSE 'Inactivo' ")
			loComandoSeleccionar.AppendLine("			END) AS Estatus ")
			loComandoSeleccionar.AppendLine("FROM		Funciones_Nomina ")
			loComandoSeleccionar.AppendLine("WHERE	Cod_Fun BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("		AND	" & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("		AND	Status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)

			Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			
			'Decodifica las fórmulas de las funciones para que se lean en "texto claro"
			Dim lnTotal As Integer = laDatosReporte.Tables(0).Rows.Count
			For i AS Integer = 0 To lnTotal - 1 
				Dim loFila As DataRow = laDatosReporte.Tables(0).Rows(i)
				
				loFila("Val_Fun") = goServicios.mDecodificarQuotedPrintable(CStr(loFila("Val_Fun")))
				
			Next i 
			
			
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rFunciones_Nomina", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrFunciones_Nomina.ReportSource = loObjetoReporte

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
