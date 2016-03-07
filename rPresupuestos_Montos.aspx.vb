
Partial Class rPresupuestos_Montos
    Inherits vis2Formularios.frmReporte
Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT		Presupuestos.Documento, ")
            loComandoSeleccionar.AppendLine("				Presupuestos.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("				Presupuestos.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("				Presupuestos.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("				Presupuestos.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("				Presupuestos.Mon_Net, ")
            loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("				Presupuestos.Status	")
            loComandoSeleccionar.AppendLine("FROM			Presupuestos, Proveedores  ")
            loComandoSeleccionar.AppendLine("WHERE			Presupuestos.Cod_Pro = Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("				AND Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("				AND Presupuestos.Fec_Ini between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("				AND Presupuestos.Cod_Pro between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("               And Presupuestos.Status     IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("				AND Presupuestos.Cod_rev between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine( " ORDER BY		Documento")


        Dim loServicios As New cusDatos.goDatos

        Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString(), "curReportes")

        loObjetoReporte	=   cusAplicacion.goReportes.mCargarReporte("rPresupuestos_Montos", laDatosReporte)
        
        Me.mTraducirReporte(loObjetoReporte)
        
		Me.mFormatearCamposReporte(loObjetoReporte)
		
        Me.crvrPresupuestos_Montos.ReportSource =	loObjetoReporte	


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
' YYG:  02/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR:  03/03/09: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' YJP:  15/05/09: Agregar filtro revisión
'-------------------------------------------------------------------------------------------'

