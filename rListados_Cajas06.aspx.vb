Imports System.Data
Partial Class rListados_Cajas06
	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
		Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

	Try

		Dim loComandoSeleccionar As New StringBuilder()

			'-------------------------------------------------------------------------------------------'
			' 1 - Select de Cobros
			'-------------------------------------------------------------------------------------------'
			loComandoSeleccionar.AppendLine(" SELECT	Cod_Caj, ")
			loComandoSeleccionar.AppendLine("			Nom_Caj, ")
			loComandoSeleccionar.AppendLine("			Status, ")
			loComandoSeleccionar.AppendLine("			Registro, ")
			loComandoSeleccionar.AppendLine("			Sal_Act, ")
			loComandoSeleccionar.AppendLine("			Sal_Efe, ")
			loComandoSeleccionar.AppendLine("			(Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Cajas ")
			loComandoSeleccionar.AppendLine(" FROM		Cajas ")
			loComandoSeleccionar.AppendLine(" WHERE		Cod_Caj		Between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("			And Status	IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" ORDER BY	Cod_Caj, ")
			loComandoSeleccionar.AppendLine("			Nom_Caj")




			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCajas", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrListados_Cajas06.ReportSource = loObjetoReporte

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
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 19/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'