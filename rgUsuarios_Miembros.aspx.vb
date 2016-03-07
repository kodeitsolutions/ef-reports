Imports System.Data
Partial Class rgUsuarios_Miembros
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
			
		
            loComandoSeleccionar.AppendLine("DECLARE @lcCliente AS CHAR(10) ")
			loComandoSeleccionar.AppendLine("SET @lcCliente = " & goServicios.mObtenerCampoFormatoSQL(goCliente.pcCodigo))
			loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Grupos.Cod_Gru					AS Cod_Gru,	")
			loComandoSeleccionar.AppendLine("			Grupos.Nom_Gru					AS Nom_Gru,	")
			loComandoSeleccionar.AppendLine("			Grupos.Nivel					AS Nivel,	")
            loComandoSeleccionar.AppendLine("			(CASE	Grupos.Status			")
			loComandoSeleccionar.AppendLine("				WHEN 'A' THEN 'Activo'		")
			loComandoSeleccionar.AppendLine("				WHEN 'I' THEN 'Inactivo'	")
            loComandoSeleccionar.AppendLine("				ELSE 'Suspendido'			")
			loComandoSeleccionar.AppendLine("			END)							AS Status_Grupo, ")
			loComandoSeleccionar.AppendLine("			Usuarios.Cod_Usu				AS Cod_Usu, ")
            loComandoSeleccionar.AppendLine("			Usuarios.Nom_Usu				AS Nom_Usu, ")
			loComandoSeleccionar.AppendLine("			(CASE	Usuarios.Status			")
            loComandoSeleccionar.AppendLine("				WHEN 'A' THEN 'Activo'		")
			loComandoSeleccionar.AppendLine("				WHEN 'I' THEN 'Inactivo'	")
			loComandoSeleccionar.AppendLine("				ELSE 'Suspendido'			")
            loComandoSeleccionar.AppendLine("			END)							AS Status_Usuario ")
			loComandoSeleccionar.AppendLine("FROM		Grupos ")
			loComandoSeleccionar.AppendLine("	JOIN	Renglones_gUsuarios ")
            loComandoSeleccionar.AppendLine("		ON	Renglones_gUsuarios.Cod_Gru = Grupos.Cod_Gru 	")
			loComandoSeleccionar.AppendLine("		AND Grupos.Cod_Gru BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("		AND Grupos.Cod_Cli = @lcCliente")
			loComandoSeleccionar.AppendLine("		AND Grupos.Sistema = "  & lcSistema)
            loComandoSeleccionar.AppendLine("	JOIN	Usuarios ")
			loComandoSeleccionar.AppendLine("		ON	Usuarios.Cod_Usu = Renglones_gUsuarios.Cod_Usu")
			loComandoSeleccionar.AppendLine("		AND Usuarios.Cod_Cli = @lcCliente")
            loComandoSeleccionar.AppendLine("WHERE		Grupos.status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")


			
			Dim loServicios As New cusDatos.goDatos
	            
			goDatos.pcNombreAplicativoExterno = "Framework"

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

				loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rgUsuarios_Miembros", laDatosReporte)
				
				Me.mTraducirReporte(loObjetoReporte)

				Me.mFormatearCamposReporte(loObjetoReporte)
		   
				Me.crvrgUsuarios_Miembros.ReportSource = loObjetoReporte
				
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
' RJG: 13/01/12: Codigo inicial
'-------------------------------------------------------------------------------------------'
