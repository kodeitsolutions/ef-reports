Imports System.Data
Partial Class rConceptos_MDetallado
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try	
	
	        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.Append(" SELECT	Cod_Con, " )
			loComandoSeleccionar.Append("			Nom_Con, " )
			loComandoSeleccionar.Append("			Status, " )
			loComandoSeleccionar.Append("			cod_isr, " )
			loComandoSeleccionar.Append("			valor1, " )
			loComandoSeleccionar.Append("			valor2, " )
			loComandoSeleccionar.Append("			valor3, " )
			loComandoSeleccionar.Append("			(Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status_Conceptos, " )
			loComandoSeleccionar.Append("			(Case When Tipo = 'E' Then 'Egreso' Else 'Ingreso' End) as Tipo " )
			loComandoSeleccionar.Append(" FROM		Conceptos " )
			loComandoSeleccionar.Append(" WHERE		Cod_Con BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.Append("			AND " & lcParametro0Hasta )
			loComandoSeleccionar.Append("			AND Status IN (" & lcParametro1Desde & ")" )
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			
		
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rConceptos_MDetallado", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrConceptos_MDetallado.ReportSource = loObjetoReporte
			
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
' GCR:  27/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
