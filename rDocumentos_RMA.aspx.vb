'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDocumentos_RMA"
'-------------------------------------------------------------------------------------------'
Partial Class rDocumentos_RMA
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try	

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			
			Dim loComandoSeleccionar As New StringBuilder()
			
			
			loComandoSeleccionar.AppendLine(" SELECT  " )
			loComandoSeleccionar.AppendLine(" 		Cod_Doc, " )
			loComandoSeleccionar.AppendLine(" 		Nom_Doc, " )
			loComandoSeleccionar.AppendLine(" 		CASE  " )
			loComandoSeleccionar.AppendLine(" 			WHEN Status = 'A' Then 'Activo' " )
			loComandoSeleccionar.AppendLine(" 			WHEN Status = 'I' Then 'Inactivo' " )
			loComandoSeleccionar.AppendLine(" 			WHEN Status = 's' Then 'Suspendido' " )
			loComandoSeleccionar.AppendLine(" 		END  AS Status, " )
			loComandoSeleccionar.AppendLine(" 		Cod_con, " )
			loComandoSeleccionar.AppendLine(" 		Tip_doc " )
			loComandoSeleccionar.AppendLine(" FROM Documentos_RMA " )
			loComandoSeleccionar.AppendLine("WHERE			Cod_Doc between " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("				And " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("				And Status IN (" & lcParametro1Desde & ")")
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			

			
			 Dim loServicios As New cusDatos.goDatos
           
           
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
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDocumentos_RMA", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrDocumentos_RMA.ReportSource = loObjetoReporte
			
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
' CMS:  11/05/10: Codigo inicial
'-------------------------------------------------------------------------------------------'