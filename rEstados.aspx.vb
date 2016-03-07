Imports System.Da
Partial Class rEstados 
     Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
		    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()


			loComandoSeleccionar.AppendLine("SELECT			Estados.Cod_est, " )
			loComandoSeleccionar.AppendLine("				Estados.Nom_est, " )
			loComandoSeleccionar.AppendLine("				Paises.Nom_pai, " )
			loComandoSeleccionar.AppendLine("				Case When Estados.Status = 'A' Then 'Activo' Else 'Inactivo' End AS Estatus " )
			loComandoSeleccionar.AppendLine("FROM			Estados, Paises " )
			
			loComandoSeleccionar.AppendLine("WHERE")
			loComandoSeleccionar.AppendLine(" 			Paises.cod_pai=Estados.cod_pai	")
		   	loComandoSeleccionar.AppendLine(" 			AND    Estados.cod_est       Between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 			AND    Paises.cod_pai       Between " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine(" 			AND		Estados.status IN (" & lcParametro2Desde & ")")
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

           Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rEstados", laDatosReporte)
			            
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrEstados.ReportSource =	 loObjetoReporte	


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
' YJP:  13/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
