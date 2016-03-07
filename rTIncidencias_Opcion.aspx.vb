'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTIncidencias_Opcion"
'-------------------------------------------------------------------------------------------'
Partial Class rTIncidencias_Opcion
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            'If lcOrdenamiento = "Igual" Then
            'loComandoSeleccionar.AppendLine("   Compras.Cod_Rev between " & lcParametro8Desde)
            'Else
            'loComandoSeleccionar.AppendLine("   Compras.Cod_Rev NOT between " & lcParametro8Desde)
            'End If

            loComandoSeleccionar.AppendLine(" SELECT    (CASE WHEN Opcion = '' THEN 'Vacio' ELSE Opcion END)    AS  Opcion, ")
            loComandoSeleccionar.AppendLine("           Can_Opc, ")
            loComandoSeleccionar.AppendLine("           Can_Err ")
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

            loComandoSeleccionar.AppendLine(" SELECT    Opcion                  AS  Opcion, ")
            loComandoSeleccionar.AppendLine("           SUM(Can_Opc)            AS  Can_Opc, ")
            loComandoSeleccionar.AppendLine("           SUM(Can_Err)            AS  Can_Err ")
            loComandoSeleccionar.AppendLine(" INTO      #curTemporal01 ")
            loComandoSeleccionar.AppendLine(" FROM      #curTemporal ")
            loComandoSeleccionar.AppendLine(" GROUP BY  Opcion ")

            loComandoSeleccionar.AppendLine(" SELECT    Opcion                  AS  Opcion, ")
            loComandoSeleccionar.AppendLine("           SUM(Can_Opc)            AS  Can_Opc, ")
            loComandoSeleccionar.AppendLine("           SUM(Can_Err)            AS  Can_Err, ")
            loComandoSeleccionar.AppendLine("           SUM(Can_Err/Can_Opc)    AS  Promedio ")
            loComandoSeleccionar.AppendLine(" INTO      #curTemporal02 ")
            loComandoSeleccionar.AppendLine(" FROM      #curTemporal01 ")
            loComandoSeleccionar.AppendLine(" GROUP BY  Opcion ")

            loComandoSeleccionar.AppendLine(" SELECT    Opcion, ")
            loComandoSeleccionar.AppendLine("           Can_Opc, ")
            loComandoSeleccionar.AppendLine("           Can_Err, ")
            loComandoSeleccionar.AppendLine("           Promedio ")
            loComandoSeleccionar.AppendLine(" FROM      #curTemporal02 ")
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTIncidencias_Opcion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTIncidencias_Opcion.ReportSource = loObjetoReporte

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
' JJD: 13/11/09: Codigo inicial
'-------------------------------------------------------------------------------------------'