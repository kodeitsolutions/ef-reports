'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMonedas"
'-------------------------------------------------------------------------------------------'
Partial Class rMonedas
     Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        

	Try	
	
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
		
            loComandoSeleccionar.AppendLine("SELECT		Cod_Mon, ")
            loComandoSeleccionar.AppendLine("			Nom_Mon, ")
            loComandoSeleccionar.AppendLine("			Status, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Status = 'A' THEN 'Activo' ELSE 'Inactivo' END) AS Status_Monedas ")
            loComandoSeleccionar.AppendLine("FROM		Monedas ")
            loComandoSeleccionar.AppendLine("WHERE		Cod_Mon BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				AND		" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND	Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

		
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")
 
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMonedas", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrMonedas.ReportSource = loObjetoReporte
			
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
' MJP:  08/07/08: Codigo inicial															'
'-------------------------------------------------------------------------------------------'
' MJP:  11/07/08: Creación objeto que cierra el archivo de reporte							'
'-------------------------------------------------------------------------------------------'
' MJP:  14/07/08: Adición filtro Status														'
'-------------------------------------------------------------------------------------------'
' MVP:  04/08/08: Cambios para multi idioma, mensaje de error y clase padre.				'
'-------------------------------------------------------------------------------------------'
' CMS:  06/05/09: Ordenamiento																'
'-------------------------------------------------------------------------------------------'
' RJG:  09/11/11: Corrección de Pie de página.												'
'-------------------------------------------------------------------------------------------'
