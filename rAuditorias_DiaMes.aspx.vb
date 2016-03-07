'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAuditorias_DiaMes"
'-------------------------------------------------------------------------------------------'
Partial Class rAuditorias_DiaMes
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    DatePart(Day, Registro) AS DiaMes,")
            loComandoSeleccionar.AppendLine(" 	   		(CASE WHEN Accion = 'Apertura'	THEN 1 ELSE 0 END) As  Apertura, ")
            loComandoSeleccionar.AppendLine(" 	   		(CASE WHEN Accion = 'Agregar'	THEN 1 ELSE 0 END) As  Agregar, ")
            loComandoSeleccionar.AppendLine(" 	   		(CASE WHEN Accion = 'Anular'	THEN 1 ELSE 0 END) As  Anular, ")
            loComandoSeleccionar.AppendLine(" 	   		(CASE WHEN Accion = 'Eliminar'	THEN 1 ELSE 0 END) As  Eliminar, ")
            loComandoSeleccionar.AppendLine(" 	   		(CASE WHEN Accion = 'Modificar'	THEN 1 ELSE 0 END) As  Modificar, ")
            loComandoSeleccionar.AppendLine(" 	   		(CASE WHEN Accion = 'Reporte'	THEN 1 ELSE 0 END) As  Reporte, ")
            loComandoSeleccionar.AppendLine(" 	   		1 As TotalOpcion ")
            loComandoSeleccionar.AppendLine(" INTO      #temporal ")
            loComandoSeleccionar.AppendLine(" FROM      Auditorias ")
            loComandoSeleccionar.AppendLine(" WHERE     Registro  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Cod_Usu     Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Tabla       Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Opcion      Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Documento   Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Codigo      Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Cod_Emp      Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)

            loComandoSeleccionar.AppendLine(" SELECT     DiaMes,")
            loComandoSeleccionar.AppendLine(" 	   		SUM(Apertura) As  Apertura, ")
            loComandoSeleccionar.AppendLine(" 	   		SUM(Agregar) As  Agregar, ")
            loComandoSeleccionar.AppendLine(" 	   		SUM(Anular) As  Anular, ")
            loComandoSeleccionar.AppendLine(" 	   		SUM(Eliminar) As  Eliminar, ")
            loComandoSeleccionar.AppendLine(" 	   		SUM(Modificar) As  Modificar, ")
            loComandoSeleccionar.AppendLine(" 	   		SUM(Reporte) As  Reporte, ")
            loComandoSeleccionar.AppendLine(" 	   		SUM(TotalOpcion) As TotalOpcion ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpOpciones ")
            loComandoSeleccionar.AppendLine(" FROM      #temporal")
            loComandoSeleccionar.AppendLine(" GROUP BY  DiaMes ")
            loComandoSeleccionar.AppendLine(" ORDER BY  DiaMes ")

            loComandoSeleccionar.AppendLine(" SELECT    DiaMes, ")
            loComandoSeleccionar.AppendLine(" 	   		Apertura, ")
            loComandoSeleccionar.AppendLine(" 	   		Agregar, ")
            loComandoSeleccionar.AppendLine(" 	   		Anular, ")
            loComandoSeleccionar.AppendLine(" 	   		Eliminar, ")
            loComandoSeleccionar.AppendLine(" 	   		Modificar, ")
            loComandoSeleccionar.AppendLine(" 	   		Reporte, ")
            loComandoSeleccionar.AppendLine(" 	   		TotalOpcion ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpOpciones ")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAuditorias_DiaMes", laDatosReporte)

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

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrAuditorias_DiaMes.ReportSource = loObjetoReporte

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
' CMS: 03/05/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 15/04/11: Ajuste de la vista de diseño
'-------------------------------------------------------------------------------------------'