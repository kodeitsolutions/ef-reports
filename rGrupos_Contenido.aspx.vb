'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rGrupos_Contenido"
'-------------------------------------------------------------------------------------------'
Partial Class rGrupos_Contenido
    Inherits vis2formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


		Try	
		
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim lcSistema As String = goServicios.mObtenerCampoFormatoSQL("Factory_" & goAplicacion.pcNombre.Trim())
            Dim loComandoSeleccionar As New StringBuilder()
			
    		
            loComandoSeleccionar.AppendLine("SELECT		Grupos.Cod_Gru, ")
            loComandoSeleccionar.AppendLine("			Grupos.Nom_Gru, ")
            loComandoSeleccionar.AppendLine("			Grupos.Nivel, ")
            loComandoSeleccionar.AppendLine("			Grupos.Sistema, ")
            loComandoSeleccionar.AppendLine("			(CASE	Grupos.Status			")
            loComandoSeleccionar.AppendLine("				WHEN 'A' THEN 'Activo'		")
			loComandoSeleccionar.AppendLine("				WHEN 'I' THEN 'Inactivo'	")
			loComandoSeleccionar.AppendLine("				WHEN 'U' Then 'Suspendido'")
			loComandoSeleccionar.AppendLine("				ELSE NULL")
			loComandoSeleccionar.AppendLine("			END)							AS Status, ")
            loComandoSeleccionar.AppendLine("			(CASE	Grupos.Tipo			")
            loComandoSeleccionar.AppendLine("				WHEN 'A' THEN 'Administrador'		")
			loComandoSeleccionar.AppendLine("				WHEN 'I' THEN 'Invitado'	")
			loComandoSeleccionar.AppendLine("				WHEN 'U' Then 'Usuario'")
			loComandoSeleccionar.AppendLine("				WHEN 'S' Then 'Supervisor'")
			loComandoSeleccionar.AppendLine("				ELSE NULL")
			loComandoSeleccionar.AppendLine("			END)							AS Tipo ")
            'loComandoSeleccionar.AppendLine("			Grupos.Accesos ")
            loComandoSeleccionar.AppendLine("FROM		Grupos ")
            loComandoSeleccionar.AppendLine("WHERE		Grupos.Cod_Gru BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Grupos.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Grupos.Cod_Cli  = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))
            loComandoSeleccionar.AppendLine("		AND Grupos.Sistema = "  & lcSistema)
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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' GCR: 01/04/09: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' MAT: 11/04/11: Ajuste de la vista de diseño.												'
'-------------------------------------------------------------------------------------------'
' RJG: 14/01/12: Estandarizacion de Código. Se agregó el filtro de Sistema.					'
'-------------------------------------------------------------------------------------------'
