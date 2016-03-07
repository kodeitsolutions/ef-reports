'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCInventarios_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rCInventarios_Fechas 
     Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Cortes.Documento, ")
            loComandoSeleccionar.AppendLine("			Cortes.Status, ")
            loComandoSeleccionar.AppendLine("			Cortes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Cortes.Comentario, ")
            loComandoSeleccionar.AppendLine("			Cortes.Contacto, ")
            loComandoSeleccionar.AppendLine("			Cortes.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			Almacenes.Nom_Alm ")
            loComandoSeleccionar.AppendLine("FROM		Cortes ")
            loComandoSeleccionar.AppendLine("	JOIN	Almacenes ON Almacenes.Cod_Alm = Cortes.Cod_Alm")
            loComandoSeleccionar.AppendLine("WHERE		Cortes.Cod_Alm		= Almacenes.Cod_Alm	")
            loComandoSeleccionar.AppendLine("		AND Cortes.Fec_Ini		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("		    AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Cortes.Documento	BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Cortes.Status		IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("      	AND Cortes.Cod_Rev		BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("      	AND Cortes.Cod_Suc		BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY    Cortes.Fec_Ini ASC, " & lcOrdenamiento) 

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCInventarios_Fechas", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
			            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCInventarios_Fechas.ReportSource = loObjetoReporte

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
' JJD: 11/10/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' YJP: 27/04/09: Agregar combo y estandarizacion
'-------------------------------------------------------------------------------------------'
' CMS: 14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP: 29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS: 15/07/09: Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' RJG: 19/01/12: Actualización de clase base: El reporte heredaba de System.UI.Page (viejo)'
'-------------------------------------------------------------------------------------------'
