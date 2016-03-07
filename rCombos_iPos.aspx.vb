'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCombos_iPos"
'-------------------------------------------------------------------------------------------'
Partial Class rCombos_iPos
Inherits vis2Formularios.frmReporte

   Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

			Try
				Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
				Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
				Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

				Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
				Dim loComandoSeleccionar As New StringBuilder()

				loComandoSeleccionar.AppendLine("SELECT		Cod_Com, ")
				loComandoSeleccionar.AppendLine("			Nom_Com, ")
				loComandoSeleccionar.AppendLine(" 			CASE")		
				loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Status = 'A' THEN 'Activo'")
				loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Status = 'I' THEN 'Inactivo'")
				loComandoSeleccionar.AppendLine(" 				WHEN Combos_iPos.Status = 'S' THEN 'Suspendido'")
				loComandoSeleccionar.AppendLine(" 			END AS Status_Combos_iPos")
				loComandoSeleccionar.AppendLine("FROM	 Combos_iPos ")
				loComandoSeleccionar.AppendLine("WHERE Cod_Com between " & lcParametro0Desde)
				loComandoSeleccionar.AppendLine(" And " & lcParametro0Hasta)
				loComandoSeleccionar.AppendLine(" And Status IN (" & lcParametro1Desde & ")")
				loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

				Dim loServicios As New cusDatos.goDatos
				Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

				loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCombos_iPos", laDatosReporte)

				Me.mTraducirReporte(loObjetoReporte)
				Me.mFormatearCamposReporte(loObjetoReporte)
				Me.crvrCombos_iPos.ReportSource = loObjetoReporte

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

				Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' MAT: 04/01/11  Codigo inicial
'-------------------------------------------------------------------------------------------'