'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rBcotizaciones_Totalizadas"
'-------------------------------------------------------------------------------------------'
Partial Class rBcotizaciones_Totalizadas
    Inherits vis2formularios.frmReporte

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

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Cotizaciones.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("           SUM((CASE WHEN Cotizaciones.Cod_For = '102' THEN Renglones_Cotizaciones.Can_Art1 ELSE 0 END)) As P_Alta, ")
            loComandoSeleccionar.AppendLine("           SUM((CASE WHEN Cotizaciones.Cod_For = '101' THEN Renglones_Cotizaciones.Can_Art1 ELSE 0 END)) As P_Media, ")
            loComandoSeleccionar.AppendLine("           SUM((CASE WHEN Cotizaciones.Cod_For = '100' THEN Renglones_Cotizaciones.Can_Art1 ELSE 0 END)) As P_Baja ")
            'loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Renglon, ")
            'loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Tra, ")
            'loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_For, ")
            'loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Cli, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Documento ")
            loComandoSeleccionar.AppendLine(" FROM      Cotizaciones, ")
            loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones ")
            loComandoSeleccionar.AppendLine(" WHERE     Cotizaciones.Documento      =   Renglones_Cotizaciones.Documento ")
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Documento  BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Fec_Ini    BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_Cli    BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Status     IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_Ven    BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cotizaciones.Cod_For    BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY  Cotizaciones.Cod_Ven ")
            'loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_For, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Renglon, ")
            'loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Tra, ")
            'loComandoSeleccionar.AppendLine("           Cotizaciones.Cod_Cli, ")
            'loComandoSeleccionar.AppendLine("           Renglones_Cotizaciones.Documento ")
            loComandoSeleccionar.AppendLine(" ORDER BY  Cotizaciones.Cod_Ven ")

            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rBcotizaciones_Totalizadas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrBcotizaciones_Totalizadas.ReportSource = loObjetoReporte

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
' MVP: 13/08/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 27/04/09: Estandarización de código
'-------------------------------------------------------------------------------------------'
' JJD: 01/05/09: Ajustes al Select
'-------------------------------------------------------------------------------------------'
