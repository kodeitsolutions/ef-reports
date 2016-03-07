Imports System.Data
Partial Class rContainers 
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			
			
			Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("SELECT	containers.tipo, " )
			loComandoSeleccionar.AppendLine("		containers.cantidad, " )
			loComandoSeleccionar.AppendLine("		containers.valor, " )
			loComandoSeleccionar.AppendLine("		containers.cod_mon, " )
			loComandoSeleccionar.AppendLine("		containers.tasa, " )
			loComandoSeleccionar.AppendLine("		CASE") 
			loComandoSeleccionar.AppendLine("			WHEN containers.Status = 'A' Then 'Activo' ")
			loComandoSeleccionar.AppendLine("			WHEN containers.Status = 'I' Then 'Inactivo'")
			loComandoSeleccionar.AppendLine("			WHEN containers.Status = 'S' Then 'Suspendido'")
			loComandoSeleccionar.AppendLine("		End AS status")	
			loComandoSeleccionar.AppendLine("FROM	containers " )
			loComandoSeleccionar.AppendLine(" WHERE	containers.tipo IN  (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine(" AND 	containers.cod_mon between " & lcParametro0Desde  )
			loComandoSeleccionar.AppendLine(" AND 	" & lcParametro0Hasta )
			loComandoSeleccionar.AppendLine(" AND 	containers.status IN (" & lcParametro2Desde & ")" )
			loComandoSeleccionar.AppendLine("ORDER BY containers."&lcOrdenamiento)
	
		    Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rContainers", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrContainers.ReportSource =	 loObjetoReporte	


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
' YJP:  24/04/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 11/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
