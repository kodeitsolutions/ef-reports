Imports System.Data
Imports cusAplicacion

Partial Class rDAuditorias_Usuarios

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
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

            loComandoSeleccionar.AppendLine(" SELECT    Cod_Usu, ")
            loComandoSeleccionar.AppendLine("           CONVERT(NCHAR(10), Registro, 103)                                                                   As	Fecha, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN DATEPART(HH, Registro) < 10 THEN '0' + CONVERT(NCHAR(2),DATEPART(HH, Registro))" _
					 & " ELSE CONVERT(NCHAR(2),DATEPART(HH, Registro)) END + ':' + CASE WHEN DATEPART(MI, Registro) < 10 THEN '0' + " _
					 & "CONVERT(NCHAR(2),DATEPART(MI, Registro)) Else CONVERT(NCHAR(2),DATEPART(MI, Registro)) END AS	Hora, ")
            loComandoSeleccionar.AppendLine("           Tipo, ")
            loComandoSeleccionar.AppendLine("           Tabla, ")
            loComandoSeleccionar.AppendLine("           Opcion, ")
            loComandoSeleccionar.AppendLine("           Accion, ")
            loComandoSeleccionar.AppendLine("           Documento, ")
            loComandoSeleccionar.AppendLine("           Codigo, ")
            loComandoSeleccionar.AppendLine("           Equipo, ")
            loComandoSeleccionar.AppendLine("           Detalle")
            loComandoSeleccionar.AppendLine(" FROM      Auditorias ")
            loComandoSeleccionar.AppendLine(" WHERE     Registro Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And       Cod_Usu       Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And       Tabla         Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And       Opcion        Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And       Documento     Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And       Codigo        Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And       Cod_Emp       Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)

            'Me.Response.Clear()
            'Me.Response.ContentType="text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return 

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------
            'If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
            '    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
            '                              "No se Encontraron Registros para los Parámetros Especificados. ", _
            '                               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
            '                               "350px", _
            '                               "200px")
            'End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDAuditorias_Usuarios", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrDAuditorias_Usuarios.ReportSource = loObjetoReporte

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
' JJD: 22/12/12: Codigo inicial
'-------------------------------------------------------------------------------------------'
