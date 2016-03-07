Imports System.Data
Partial Class rGrupos_Usuarios
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
			loComandoSeleccionar.AppendLine("			Grupos.Status, ")
			loComandoSeleccionar.AppendLine("			Grupos.Nivel, ")
			loComandoSeleccionar.AppendLine("			(CASE	Grupos.Status			")
			loComandoSeleccionar.AppendLine("				WHEN 'A' THEN 'Activo'		")
			loComandoSeleccionar.AppendLine("				WHEN 'I' THEN 'Inactivo'	")
			loComandoSeleccionar.AppendLine("				WHEN 'U' Then 'Suspendido'	")
			loComandoSeleccionar.AppendLine("				ELSE NULL")
			loComandoSeleccionar.AppendLine("			END)			AS Status_Grupos_Usuarios")
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

				loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rGrupos_Usuarios", laDatosReporte)
				
				Me.mTraducirReporte(loObjetoReporte)

				Me.mFormatearCamposReporte(loObjetoReporte)
		   
				Me.crvrGrupos_Usuarios.ReportSource = loObjetoReporte
				
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
' MJP   :  09/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP :  11/07/08 : Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MJP :  14/07/08 : Agregacion filtro Status
'-------------------------------------------------------------------------------------------'
' GCR :  26/03/09 : Estandarizacion de codigo y adicion de parametros   y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  31/08/09: Metodo de ordenamiento, verificacionde registros, Boton Imprimir
'-------------------------------------------------------------------------------------------'
' MAT: 11/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'
' RJG: 14/01/12: Estandarizacion de Código. Se agregó el filtro de Sistema.					'
'-------------------------------------------------------------------------------------------'
