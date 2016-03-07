'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_Valor"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_Valor
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Uni1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Act1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cos_Ult1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cos_Ult2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cos_Pro1, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cos_Pro2, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Dep, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Sec, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Cla, ")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Mar ")
            loComandoSeleccionar.AppendLine("FROM  		Articulos ")
            loComandoSeleccionar.AppendLine("WHERE 		Articulos.Cod_Art       BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Articulos.Status        IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Dep       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Sec       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Mar       BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Tip       BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Cla       BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("		AND Articulos.Cod_Ubi    BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro7Hasta)
			loComandoSeleccionar.AppendLine("		AND	Articulos.Exi_Act1 <> 0 " )
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_Valor", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_Valor.ReportSource = loObjetoReporte


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
' MVP: 09/07/08: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
' MVP: 01/08/08: Cambios para multi idioma, mensaje de error y clase padre.					'
'-------------------------------------------------------------------------------------------'
' JFP: 09/10/08: Limitacion de la longitud del nombre y ajustes de forma de presentacion.	'
'-------------------------------------------------------------------------------------------'
' JFP: 25/10/08: Adecuacion al nuevo tipo de reporte.										'
'-------------------------------------------------------------------------------------------'
' JJD: 02/02/09: Adecuacion al nuevo tipo de reporte.										'
'-------------------------------------------------------------------------------------------'
' JJD: 24/02/09: Aplicacion de filtro a las Secciones de los Articulos para no duplicarlos.	'
'-------------------------------------------------------------------------------------------'
' CMS: 05/05/09: Ordenamiento																'
'-------------------------------------------------------------------------------------------'
' CMS: 11/08/09: Verificacion de registros y se agregaro el filtro_ Ubicación				'
'-------------------------------------------------------------------------------------------'
' RJG: 23/03/12: Se agregó un filtro fijo para solo mostrar con existencia <> 0.			'
'-------------------------------------------------------------------------------------------'
