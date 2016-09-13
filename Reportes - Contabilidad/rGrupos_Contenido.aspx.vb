Imports System.Data
Partial Class rGrupos_Contenido
    Inherits vis2formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


		Try	
		
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
			
    		
            loComandoSeleccionar.AppendLine("SELECT		Cod_Gru, ")
            loComandoSeleccionar.AppendLine("				Nom_Gru, ")
            loComandoSeleccionar.AppendLine("				Nivel, ")
            loComandoSeleccionar.AppendLine("				Sistema, ")
            loComandoSeleccionar.AppendLine("				(Case When Status = 'A' Then 'Activo' Else 'Inactivo' End) as Status, ")
            loComandoSeleccionar.AppendLine("				(Case Tipo When 'A' Then 'Administradores' When 'U' Then 'Usuarios' When 'S' Then 'Supervisores 'Else 'Invitado' End) as Tipo, ")
            loComandoSeleccionar.AppendLine("				Accesos ")
            loComandoSeleccionar.AppendLine("				FROM	 Grupos ")
            loComandoSeleccionar.AppendLine("WHERE			Cod_Gru between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("				AND status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("		    AND Cod_Cli  = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

		
				Dim loServicios As New cusDatos.goDatos
	            
				goDatos.pcNombreAplicativoExterno = "Framework"

				Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
				
				loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rGrupos_Contenido", laDatosReporte)
		   
				 Me.mTraducirReporte(loObjetoReporte)

				Me.mFormatearCamposReporte(loObjetoReporte)
				
				Me.crvrGrupos_Contenido.ReportSource = loObjetoReporte
				
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
' GCR:  01/04/09 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 11/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
