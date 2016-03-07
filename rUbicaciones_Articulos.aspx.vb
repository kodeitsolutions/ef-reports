'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rUbicaciones_Articulos"
'-------------------------------------------------------------------------------------------'
Partial Class rUbicaciones_Articulos
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        
		Try
		
		    Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
		
			loComandoSeleccionar.AppendLine("SELECT		Cod_Ubi, " )
			loComandoSeleccionar.AppendLine("			Nom_Ubi, " )
			loComandoSeleccionar.AppendLine("			Case When Status = 'A' Then 'Activo' Else 'Inactivo' End as Status " )
			loComandoSeleccionar.AppendLine("FROM		Ubicaciones_Articulos " )
			loComandoSeleccionar.AppendLine("WHERE		Cod_Ubi BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("			AND Status IN (" & lcParametro1Desde & ")")
            'loComandoSeleccionar.AppendLine("ORDER BY	Cod_Ubi ") 
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
  
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rUbicaciones_Articulos", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)
			
			Me.crvrUbicaciones_Articulos.ReportSource = loObjetoReporte

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
' JJD: 04/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR: 11/03/09: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  06/05/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'