Imports System.Data
Partial Class rBancos
   Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        

	Try	
	
		Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
		Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
		Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
		Dim loComandoSeleccionar As New StringBuilder() 

		loComandoSeleccionar.AppendLine("SELECT		Cod_Ban, " )
		loComandoSeleccionar.AppendLine("			Nom_Ban, " )
		loComandoSeleccionar.AppendLine("			CASE")
		loComandoSeleccionar.AppendLine("				WHEN Bancos.Status = 'A' Then 'Activo'	")
		loComandoSeleccionar.AppendLine("				WHEN Bancos.Status = 'I' Then 'Inactivo'	")
		loComandoSeleccionar.AppendLine("				WHEN Bancos.Status = 'S' Then 'Suspendido'	")
		loComandoSeleccionar.AppendLine("			END AS Status" )
		loComandoSeleccionar.AppendLine("FROM		Bancos " )
		loComandoSeleccionar.AppendLine("	WHERE	Cod_Ban between " & lcParametro0Desde )
		loComandoSeleccionar.AppendLine("	AND		" & lcParametro0Hasta )
		loComandoSeleccionar.AppendLine("	AND		Status IN (" & lcParametro1Desde & ")" )
		loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)



		
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rBancos", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
            Me.crvrBancos.ReportSource = loObjetoReporte
	   

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
' MVP:  01/08/08: Cambios para multi idioma, mensaje de error y clase padre.
'-------------------------------------------------------------------------------------------'
' YJP:  22/04/09: Agregar combo a estatus y estandarizacion de codigo.
'-------------------------------------------------------------------------------------------'
' MAT:  19/04/11 : Ajuste de la vista de diseño.
'-------------------------------------------------------------------------------------------'