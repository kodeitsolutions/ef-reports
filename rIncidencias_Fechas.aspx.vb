'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rIncidencias_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rIncidencias_Fechas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))

            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))

            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

            If lcParametro3Desde = "''" Then
                lcParametro3Hasta = "'zzzzzzz'"
            Else
                lcParametro3Hasta = lcParametro3Desde
            End If

            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            If lcParametro4Desde = "''" Then
                lcParametro4Hasta = "'zzzzzzz'"
            Else
                lcParametro4Hasta = lcParametro4Desde
            End If

            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            If lcParametro5Desde = "''" Then
                lcParametro5Hasta = "'zzzzzzz'"
            Else
                lcParametro5Hasta = lcParametro5Desde
            End If

            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    CONVERT(NCHAR(10), Registro, 103)                                       AS  Fecha,	")
            loComandoSeleccionar.AppendLine(" 			CONVERT(NCHAR(10), Registro, 112)                                       AS  Fecha2,	")
            loComandoSeleccionar.AppendLine(" 			Cod_Usu                                                                 AS  Cod_Usu, ")
            loComandoSeleccionar.AppendLine("           Cod_Emp                                                                 AS  Cod_Emp, ")
            loComandoSeleccionar.AppendLine("           Sistema                                                                 AS  Sistema, ")
            loComandoSeleccionar.AppendLine("           Modulo                                                                  AS  Modulo, ")
            loComandoSeleccionar.AppendLine("           Opcion                                                                  AS  Opcion, ")
            loComandoSeleccionar.AppendLine("           Can_Err                                                                 AS  Can_Err, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN " & lcParametro6Desde & " = 'Si' THEN Detalles ELSE '' END    AS  Mensaje ")
            loComandoSeleccionar.AppendLine(" INTO      #curTemporal ")
            loComandoSeleccionar.AppendLine(" FROM      Factory_Global.dbo.Errores ")
            loComandoSeleccionar.AppendLine(" WHERE     Registro	    BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Usu     BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Status      IN ( " & lcParametro2Desde & " )")
            loComandoSeleccionar.AppendLine("           AND Sistema     BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Modulo      BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Opcion      BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
 


            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return


            loComandoSeleccionar.AppendLine(" SELECT    Fecha,	")
            loComandoSeleccionar.AppendLine(" 			Fecha2,	")
            loComandoSeleccionar.AppendLine(" 			Cod_Usu, ")
            loComandoSeleccionar.AppendLine("           Cod_Emp, ")
            loComandoSeleccionar.AppendLine("           Sistema, ")
            loComandoSeleccionar.AppendLine("           Modulo, ")
            loComandoSeleccionar.AppendLine("           Opcion, ")
            loComandoSeleccionar.AppendLine("           Can_Err, ")
            loComandoSeleccionar.AppendLine("           Mensaje ")
            loComandoSeleccionar.AppendLine(" FROM      #curTemporal  ")
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)




            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rIncidencias_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrIncidencias_Fechas.ReportSource = loObjetoReporte

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
' JJD: 23/12/09: Codigo inicial
'-------------------------------------------------------------------------------------------'