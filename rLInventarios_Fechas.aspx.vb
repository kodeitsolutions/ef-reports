'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLInventarios_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rLInventarios_Fechas 
    Inherits vis2formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden


	Try
	
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

			 Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("  SELECT			Libres_Inventarios.Documento, ")
            loComandoSeleccionar.AppendLine(" 					Libres_Inventarios.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 					Libres_Inventarios.Comentario, ")
            loComandoSeleccionar.AppendLine(" 					Libres_Inventarios.Status, ")
            loComandoSeleccionar.AppendLine(" 					Libres_Inventarios.Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 					Libres_Inventarios.Mon_Bru, ")
            loComandoSeleccionar.AppendLine(" 					Libres_Inventarios.Mon_Imp, ")
            loComandoSeleccionar.AppendLine(" 					Libres_Inventarios.Mon_Net ")
            loComandoSeleccionar.AppendLine(" FROM				Libres_Inventarios ")
            loComandoSeleccionar.AppendLine(" WHERE				Libres_Inventarios.Fec_Ini		 BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 					AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 					AND Libres_Inventarios.Documento BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 					AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 					AND Libres_Inventarios.Status	 IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("                   AND Libres_Inventarios.Cod_Rev between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 		            AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY     CONVERT(nchar(30), Libres_Inventarios.Fec_Ini,112), " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
            
           ' Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rLInventarios_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)
            
            Me.crvrLInventarios_Fechas.ReportSource =	 loObjetoReporte	

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
' GMO: 07/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR: 10/03/08: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  02/07/09 : Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'