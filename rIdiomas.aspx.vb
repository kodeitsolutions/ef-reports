'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rIdiomas"
'-------------------------------------------------------------------------------------------'
Partial Class rIdiomas
Inherits vis2Formularios.frmReporte

   Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

			Try	
				
				Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
				Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
				Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
				
				Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
				Dim loComandoSeleccionar As New StringBuilder()
				
				loComandoSeleccionar.AppendLine(" SELECT")				
				loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Cod_Idi AS Cod_Idi,")
				loComandoSeleccionar.AppendLine(" 			Idiomas.Nom_Idi AS Nom_Idi,")
				loComandoSeleccionar.AppendLine(" 			Renglones_Idiomas.Comentario,")
				loComandoSeleccionar.AppendLine(" 			CASE")		
				loComandoSeleccionar.AppendLine(" 				WHEN Renglones_Idiomas.Tip_Obj = 'lbl' THEN 'Etiqueta'")
				loComandoSeleccionar.AppendLine(" 				WHEN Renglones_Idiomas.Tip_Obj = 'opc' THEN 'Opcion'")
				loComandoSeleccionar.AppendLine(" 				WHEN Renglones_Idiomas.Tip_Obj = 'chk' THEN 'Verificacion'")
				loComandoSeleccionar.AppendLine(" 				WHEN Renglones_Idiomas.Tip_Obj = 'frm' THEN 'Formulario'")
				loComandoSeleccionar.AppendLine(" 				WHEN Renglones_Idiomas.Tip_Obj = 'cmd' THEN 'Boton'")
				loComandoSeleccionar.AppendLine(" 			END AS Tip_Obj")
				loComandoSeleccionar.AppendLine(" FROM Renglones_Idiomas")
				loComandoSeleccionar.AppendLine(" LEFT JOIN Idiomas ON (Renglones_Idiomas.Cod_Idi = Idiomas.Cod_Idi) ") 			
				loComandoSeleccionar.AppendLine(" WHERE Renglones_Idiomas.Cod_Idi between " & lcParametro0Desde)
				loComandoSeleccionar.AppendLine(" AND " & lcParametro0Hasta)
				loComandoSeleccionar.AppendLine(" And Tip_Obj IN (" & lcParametro1Desde & ")")
				loComandoSeleccionar.AppendLine(" ORDER BY      " & lcOrdenamiento)
				
				goDatos.pcNombreAplicativoExterno = "Framework"

				'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
				
				Dim loServicios As New cusDatos.goDatos
				Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

				loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rIdiomas", laDatosReporte)

				Me.mTraducirReporte(loObjetoReporte)
				Me.mFormatearCamposReporte(loObjetoReporte)
				Me.crvrIdiomas.ReportSource = loObjetoReporte

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
' MAT: 11/01/11  Codigo inicial
'-------------------------------------------------------------------------------------------'