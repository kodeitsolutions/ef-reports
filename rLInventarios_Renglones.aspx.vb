'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLInventarios_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rLInventarios_Renglones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
	
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

	
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Libres_Inventarios.Documento, ")
            loComandoSeleccionar.AppendLine("			Libres_Inventarios.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Libres_Inventarios.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Libres_Inventarios.Status, ")
            loComandoSeleccionar.AppendLine("			Renglones_LInventarios.Renglon, ")
            loComandoSeleccionar.AppendLine("			Renglones_LInventarios.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Renglones_LInventarios.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			Renglones_LInventarios.Can_Art1, ")
            loComandoSeleccionar.AppendLine("			Renglones_LInventarios.Precio1, ")
            loComandoSeleccionar.AppendLine("			(Renglones_LInventarios.Can_Art1 * Renglones_LInventarios.Precio1) AS SubTotal, ")
            loComandoSeleccionar.AppendLine("			Renglones_LInventarios.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			((Renglones_LInventarios.Can_Art1 * Renglones_LInventarios.Precio1) + Renglones_LInventarios.Mon_Imp1) AS Neto, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art ")
            loComandoSeleccionar.AppendLine(" FROM		Libres_Inventarios, ")
            loComandoSeleccionar.AppendLine("			Renglones_LInventarios, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Libres_Inventarios.Documento		=	Renglones_LInventarios.Documento ")
            loComandoSeleccionar.AppendLine("			And Renglones_LInventarios.Cod_Art	=	Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("			And Libres_Inventarios.Documento	Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Libres_Inventarios.Fec_Ini		Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Libres_Inventarios.Status		IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Libres_Inventarios.Cod_Rev between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 		    AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY	Libres_Inventarios.Documento, Renglones_LInventarios.Renglon ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rLInventarios_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLInventarios_Renglones.ReportSource =	 loObjetoReporte	

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
' GCR: 10/03/09: Estandarizacion de codigo y ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'