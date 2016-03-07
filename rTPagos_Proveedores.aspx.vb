'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTPagos_Proveedores"
'-------------------------------------------------------------------------------------------'
Partial Class rTPagos_Proveedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT    Cuentas_Pagar.Cod_Pro                 AS Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(Cuentas_Pagar.Mon_Net,0))  AS Mon_Net ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTemporalCuentas ")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Pagar.Cod_Tip     =    'Fact'")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Fec_Ini     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Pro     BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Ven     BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Status      IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Mon     BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Rev     BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Suc     BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY  Cuentas_Pagar.Cod_Pro ")

            loComandoSeleccionar.AppendLine(" SELECT    Pagos.Cod_Pro                                                                                    AS Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Pagos.Tip_Ope = 'Efectivo'       THEN Detalles_Pagos.Mon_Net ELSE 0 END) AS Efectivo, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Pagos.Tip_Ope = 'Ticket'         THEN Detalles_Pagos.Mon_Net ELSE 0 END) AS Ticket, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Pagos.Tip_Ope = 'Cheque'         THEN Detalles_Pagos.Mon_Net ELSE 0 END) AS Cheque, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Pagos.Tip_Ope = 'Tarjeta'        THEN Detalles_Pagos.Mon_Net ELSE 0 END) AS Tarjeta, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Pagos.Tip_Ope = 'Deposito'       THEN Detalles_Pagos.Mon_Net ELSE 0 END) AS Deposito, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Pagos.Tip_Ope = 'Transferencia'  THEN Detalles_Pagos.Mon_Net ELSE 0 END) AS Transferencia ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTemporal ")
            loComandoSeleccionar.AppendLine(" FROM      Pagos JOIN Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine(" WHERE     Pagos.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Pro  BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Ven  BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Status   IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Mon  BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Rev  BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Suc  BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY  Pagos.Cod_Pro ")

            loComandoSeleccionar.AppendLine(" SELECT    ISNULL(#tmpTemporal.Cod_Pro, #tmpTemporalCuentas.Cod_Pro)	AS  Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(#tmpTemporalCuentas.Mon_Net,0))                  AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(#tmpTemporal.Efectivo,0))                        AS  Efectivo, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(#tmpTemporal.Ticket,0))                          AS  Ticket, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(#tmpTemporal.Cheque,0))                          AS  Cheque, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(#tmpTemporal.Tarjeta,0))                         AS  Tarjeta, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(#tmpTemporal.Deposito,0))                        AS  Deposito, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(#tmpTemporal.Transferencia,0))                   AS  Transferencia, ")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(#tmpTemporal.Efectivo,0) + ISNULL(#tmpTemporal.Cheque,0) + ISNULL(#tmpTemporal.Tarjeta,0) + ISNULL(#tmpTemporal.Deposito,0) + ISNULL(#tmpTemporal.Transferencia,0) + ISNULL(#tmpTemporal.Ticket,0))  AS Total_Pagos ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTemporal001 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTemporal FULL JOIN  #tmpTemporalCuentas ON (#tmpTemporal.Cod_Pro = #tmpTemporalCuentas.Cod_Pro) ")
            loComandoSeleccionar.AppendLine(" GROUP BY  #tmpTemporal.Cod_Pro, #tmpTemporalCuentas.Cod_Pro ")

            loComandoSeleccionar.AppendLine(" SELECT    #tmpTemporal001.Cod_Pro         AS  Cod_Pro, ")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro             AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           #tmpTemporal001.Mon_Net         AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           #tmpTemporal001.Efectivo        AS  Efectivo, ")
            loComandoSeleccionar.AppendLine("           #tmpTemporal001.Ticket          AS  Ticket, ")
            loComandoSeleccionar.AppendLine("           #tmpTemporal001.Cheque          AS  Cheque, ")
            loComandoSeleccionar.AppendLine("           #tmpTemporal001.Tarjeta         AS  Tarjeta, ")
            loComandoSeleccionar.AppendLine("           #tmpTemporal001.Deposito        AS  Deposito, ")
            loComandoSeleccionar.AppendLine("           #tmpTemporal001.Transferencia   AS  Transferencia, ")
            loComandoSeleccionar.AppendLine("           #tmpTemporal001.Total_Pagos     AS  Total_Pagos ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTemporal001, Proveedores ")
            loComandoSeleccionar.AppendLine(" WHERE     #tmpTemporal001.Cod_Pro    =   Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine(" ORDER BY  #tmpTemporal001." & lcOrdenamiento)

            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTPagos_Proveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTPagos_Proveedores.ReportSource = loObjetoReporte

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
' JJD: 07/01/10: Programacion inicial
'-------------------------------------------------------------------------------------------'
