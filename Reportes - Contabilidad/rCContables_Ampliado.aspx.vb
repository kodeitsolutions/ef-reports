'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCContables_Ampliado"
'-------------------------------------------------------------------------------------------'
Partial Class rCContables_Ampliado
     Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
		
		Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
	    
		Dim loComandoSeleccionar As New StringBuilder()
		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		Try	
			
			loComandoSeleccionar.AppendLine("SELECT		Cuentas_Contables.Cod_Cue,")
			loComandoSeleccionar.AppendLine("			SPACE(LEN(RTRIM(Cuentas_Contables.Cod_Cue))) + Cuentas_Contables.Nom_Cue AS Nom_Cue,")
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Status ,")
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Cod_Cen,")
			loComandoSeleccionar.AppendLine("			Centros_Costos.Nom_Cen,")
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Mon_Ini,")
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Auxiliar,")
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Gasto, ")
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Movimiento ")
			loComandoSeleccionar.AppendLine("FROM		Cuentas_Contables ")
			loComandoSeleccionar.AppendLine("	JOIN	Centros_Costos")
			loComandoSeleccionar.AppendLine("		ON	Centros_Costos.Cod_Cen = Cuentas_Contables.Cod_Cen")
			loComandoSeleccionar.AppendLine("WHERE		Cuentas_Contables.Cod_Cue BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("			AND Cuentas_Contables.Status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine("ORDER BY	Cuentas_Contables." & lcOrdenamiento)

			
			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString(), "curReportes")

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCContables_Ampliado", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrCContables_Ampliado.ReportSource = loObjetoReporte
	   

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
' RJG: 10/01/12: Codigo inicial
'-------------------------------------------------------------------------------------------'
