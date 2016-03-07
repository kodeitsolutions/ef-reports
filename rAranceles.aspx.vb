Imports System.Data
Partial Class rAranceles
    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
		
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("SELECT	aranceles.cod_ara, " )
			loComandoSeleccionar.AppendLine("		aranceles.nom_ara, " )
			loComandoSeleccionar.AppendLine("		aranceles.tipo, " )
			loComandoSeleccionar.AppendLine("		aranceles.tip_ara, " )
			loComandoSeleccionar.AppendLine("		aranceles.status, ")
			loComandoSeleccionar.AppendLine("Case When Status = 'A' Then 'Activo' Else 'Inactivo' End as Status_aranceles " )
						
			loComandoSeleccionar.AppendLine("FROM	aranceles " )
			
			loComandoSeleccionar.AppendLine(" WHERE	aranceles.cod_ara between " & lcParametro0Desde  )
			loComandoSeleccionar.AppendLine(" AND 	" & lcParametro0Hasta )
			loComandoSeleccionar.AppendLine(" AND 	aranceles.status IN (" & lcParametro1Desde & ")" )
			loComandoSeleccionar.AppendLine("ORDER BY      aranceles."&lcOrdenamiento)
					
		    Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rAranceles", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrAranceles.ReportSource =	 loObjetoReporte	


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
' MAT: 18/04/11: Mejora en la vista de Diseño
'-------------------------------------------------------------------------------------------'
