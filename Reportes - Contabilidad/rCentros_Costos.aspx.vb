Imports System.Data
Partial Class rCentros_Costos
    Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
	    
		Dim loComandoSeleccionar As New StringBuilder()
		
		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

	Try	
		
		loComandoSeleccionar.AppendLine("SELECT		Centros_Costos.Cod_Cen, " )
		loComandoSeleccionar.AppendLine("			Centros_Costos.Nom_Cen, " )
		loComandoSeleccionar.AppendLine("			Centros_Costos.Status, " )
		loComandoSeleccionar.AppendLine("			(Case Centros_Costos.Status When   'A' Then 'Activo' When 'S' then 'Suspendido' Else 'Inactivo' End) as Status_Centros_Costos " )
		loComandoSeleccionar.AppendLine("FROM		Centros_Costos " )
		loComandoSeleccionar.AppendLine("WHERE		Centros_Costos.Cod_Cen between " & lcParametro0Desde)
		loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
		loComandoSeleccionar.AppendLine("			AND Centros_Costos.Status IN (" & lcParametro1Desde & ")")
		loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)

		Dim loServicios As New cusDatos.goDatos

		Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString(), "curReportes")

		loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCentros_Costos", laDatosReporte)

		Me.mTraducirReporte(loObjetoReporte)
		        
		Me.mFormatearCamposReporte(loObjetoReporte)


		Me.crvrCentros_Costos.ReportSource = loObjetoReporte
	   

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
' MJP   :  15/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------' 
' MVP:  05/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' MAT:  16/05/11: Mejora de la vista de diseño, Ajuste del Select
'-------------------------------------------------------------------------------------------' 
