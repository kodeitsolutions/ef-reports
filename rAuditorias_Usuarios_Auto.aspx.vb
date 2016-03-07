Imports System.Data
Imports cusAplicacion
Partial Class rAuditorias_Usuarios_Auto
    Inherits vis2formularios.frmReporteAutomatico
    'Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Try

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
        Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
        Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
        Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
        Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
        Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
        Dim loComandoSeleccionar As New StringBuilder()


        loComandoSeleccionar.AppendLine(" SELECT Cod_Usu, ")
        loComandoSeleccionar.AppendLine(" 			SUM((CASE WHEN Accion = 'Apertura'	THEN 1 ELSE 0 END)) As Apertura, ")
        loComandoSeleccionar.AppendLine(" 			SUM((CASE WHEN Accion = 'Agregar'	    THEN 1 ELSE 0 END)) As Agregar, ")
        loComandoSeleccionar.AppendLine(" 			SUM((CASE WHEN Accion = 'Anular'		THEN 1 ELSE 0 END)) As Anular, ")
        loComandoSeleccionar.AppendLine(" 			SUM((CASE WHEN Accion = 'Eliminar'	THEN 1 ELSE 0 END)) As Eliminar, ")
        loComandoSeleccionar.AppendLine(" 			SUM((CASE WHEN Accion = 'Modificar' OR Accion = 'Confirmar' OR Accion = 'Reversar'    THEN 1 ELSE 0 END)) As Modificar, ")
        loComandoSeleccionar.AppendLine(" 			SUM((CASE WHEN Accion = 'Reporte'	OR Accion = 'Formato'	THEN 1 ELSE 0 END)) As Reporte, ")
        loComandoSeleccionar.AppendLine(" 			SUM((CASE WHEN Accion = ''	THEN 1 ELSE 0 END)) As Vacio, ")
        loComandoSeleccionar.AppendLine(" 			SUM((CASE WHEN Accion = 'Importar'	THEN 1 ELSE 0 END)) As Importar, ")
        loComandoSeleccionar.AppendLine(" 			SUM((CASE WHEN Accion = 'Procesar'	THEN 1 ELSE 0 END)) As Procesar, ")
        loComandoSeleccionar.AppendLine(" 			SUM(1) As TotalReporte ")
        loComandoSeleccionar.AppendLine(" FROM Auditorias ")
        loComandoSeleccionar.AppendLine(" WHERE Registro Between " & lcParametro0Desde)
        loComandoSeleccionar.AppendLine(" 			And " & lcParametro0Hasta)
        loComandoSeleccionar.AppendLine(" 			And Cod_Usu Between " & lcParametro1Desde)
        loComandoSeleccionar.AppendLine(" 			And " & lcParametro1Hasta)
        loComandoSeleccionar.AppendLine(" 			And Tabla Between " & lcParametro2Desde)
        loComandoSeleccionar.AppendLine(" 			And " & lcParametro2Hasta)
        loComandoSeleccionar.AppendLine(" 			And Opcion Between " & lcParametro3Desde)
        loComandoSeleccionar.AppendLine(" 			And " & lcParametro3Hasta)
        loComandoSeleccionar.AppendLine(" 			And Documento Between " & lcParametro4Desde)
        loComandoSeleccionar.AppendLine(" 			And " & lcParametro4Hasta)
        loComandoSeleccionar.AppendLine(" 			And Codigo Between " & lcParametro5Desde)
        loComandoSeleccionar.AppendLine(" 			And " & lcParametro5Hasta)
        loComandoSeleccionar.AppendLine(" 			And Cod_Emp Between " & lcParametro6Desde)
        loComandoSeleccionar.AppendLine(" 			And " & lcParametro6Hasta)
        loComandoSeleccionar.AppendLine(" GROUP BY  Auditorias.Cod_Usu ")
        loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento)

        Dim loServicios As New cusDatos.goDatos

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

        loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAuditorias_Usuarios_Auto", laDatosReporte)

        Me.mTraducirReporte(loObjetoReporte)

        Me.mFormatearCamposReporte(loObjetoReporte)

        Me.crvrAuditorias_Usuarios_Auto.ReportSource = loObjetoReporte

        'Catch loExcepcion As Exception
        '
        '   Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
        '                 "No se pudo Completar el Proceso: " & loExcepcion.Message, _
        '                 vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
        '                "auto", _
        '               "auto")
        '
        'End Try

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
' EAG: 31/07/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
