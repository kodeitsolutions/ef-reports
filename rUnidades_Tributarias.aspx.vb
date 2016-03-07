Imports System.Data
Partial Class rUnidades_Tributarias
     Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try	
		Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
		Dim loComandoSeleccionar As New StringBuilder()      	
	
		loComandoSeleccionar.AppendLine("SELECT	Cod_Uni, " )
		loComandoSeleccionar.AppendLine("Fec_Gac, " )
		loComandoSeleccionar.AppendLine("Monto " )
		loComandoSeleccionar.AppendLine("FROM	 Unidades_Tributarias " )
		loComandoSeleccionar.AppendLine("WHERE Cod_Uni between " & lcParametro0Desde )
		loComandoSeleccionar.AppendLine(" And " & lcParametro0Hasta )
		loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


		
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rUnidades_Tributarias", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrUnidades_Tributarias.ReportSource = loObjetoReporte
			
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
' MJP   :  08/07/08 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' MJP :  11/07/08 : Creación objeto que cierra el archivo de reporte
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' CMS:  07/08/09: Verificacion de registros
'-------------------------------------------------------------------------------------------'