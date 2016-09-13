Imports System.Data
Partial Class rCuentas_Integracion
     Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
	    
	    Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
	    
		Dim loComandoSeleccionar As New StringBuilder()

	Try	
	
			loComandoSeleccionar.AppendLine("SELECT		Cuentas_Integracion.Cuenta, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Integracion.Nombre, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Integracion.Status, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Integracion.Cod_Cue, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Nom_Cue, " )
			loComandoSeleccionar.AppendLine("			(Case Cuentas_Integracion.Status When   'A' Then 'Activo' When 'S' then 'Suspendido' Else 'Inactivo' End) as Status_Cuentas_Integracion " )
			loComandoSeleccionar.AppendLine("FROM		Cuentas_Integracion, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Contables " )
			loComandoSeleccionar.AppendLine("WHERE		Cuentas_Integracion.Cod_Cue = Cuentas_Contables.Cod_Cue")
			loComandoSeleccionar.AppendLine("			AND Cuentas_Integracion.Cuenta between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("			AND Cuentas_Integracion.Status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)

		
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCuentas_Integracion", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCuentas_Integracion.ReportSource = loObjetoReporte
	   

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
' GCR:  02/04/09: Estandarización de codigo y adicion de campo cuenta contable.
'-------------------------------------------------------------------------------------------' 
' MAT:  16/05/11: Mejora de la vista de diseño
'-------------------------------------------------------------------------------------------' 
