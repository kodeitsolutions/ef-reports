'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCInventarios_Renglones"
'-------------------------------------------------------------------------------------------'
Partial Class rCInventarios_Renglones
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

	Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            
             Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

		Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Cortes.Documento, ")
            loComandoSeleccionar.AppendLine("			Cortes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Cortes.Status, ")
            loComandoSeleccionar.AppendLine("			Cortes.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("			Almacenes.Nom_Alm, ")
            loComandoSeleccionar.AppendLine("			Cortes.Comentario, ")
            loComandoSeleccionar.AppendLine("			Cortes.Contacto, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cortes.Renglon, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cortes.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cortes.Can_Teo, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cortes.Can_Rea, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cortes.Cos_Pro1 AS Costo, ")
            loComandoSeleccionar.AppendLine("			(Renglones_Cortes.Can_Rea * Renglones_Cortes.Costo) AS	Cos_Aju ")
            loComandoSeleccionar.AppendLine(" FROM		Cortes, ")
            loComandoSeleccionar.AppendLine("			Renglones_Cortes, ")
            loComandoSeleccionar.AppendLine("			Almacenes, ")
            loComandoSeleccionar.AppendLine("			Articulos ")
            loComandoSeleccionar.AppendLine(" WHERE		Cortes.Documento					=	Renglones_Cortes.Documento ")
            loComandoSeleccionar.AppendLine("			And Renglones_Cortes.Cod_Art		=	Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("			And Cortes.Cod_Alm					=	Almacenes.Cod_Alm ")
            loComandoSeleccionar.AppendLine("			And Cortes.Documento				Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			And Cortes.Fec_Ini					Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND		 " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			And Articulos.Cod_Art				Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND		 " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			And Cortes.Status					IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("      	      AND Cortes.Cod_Rev between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 		      AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("      	      AND Cortes.Cod_Suc between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 		      AND " & lcParametro5Hasta)
            'loComandoSeleccionar.AppendLine(" ORDER BY	Cortes.Documento, Renglones_Cortes.Renglon ")
             loComandoSeleccionar.AppendLine("ORDER BY  Cortes.Documento, " & lcOrdenamiento)

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

            
            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rCInventarios_Renglones", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCInventarios_Renglones.ReportSource =	 loObjetoReporte	

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
' YJP: 27/04/09: Agragar combo estatus y estandarizacion
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  29/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  08/04/10: Metodo de ordenamiento, verificacionde registros
'-------------------------------------------------------------------------------------------'
' CMS:  08/04/10: Se cambio el origen de la columna costo de renglones costo a cos_pro1
'-------------------------------------------------------------------------------------------'