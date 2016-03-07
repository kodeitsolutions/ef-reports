'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTCobros_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rTCobros_Fechas
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
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(YEAR, Fec_Ini)     AS Año,")
            loComandoSeleccionar.AppendLine("           DATEPART(MONTH, Fec_Ini)    AS Mes,")
            loComandoSeleccionar.AppendLine("           DATEPART(DAY, Fec_Ini)      AS Dia,")
            loComandoSeleccionar.AppendLine("           SUM(ISNULL(Mon_Net,0))      AS Mon_Net")
            loComandoSeleccionar.AppendLine(" INTO      #Temporal")
            loComandoSeleccionar.AppendLine(" FROM      Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine(" WHERE     Cuentas_Cobrar.Cod_Tip     IN ('Fact','ATD')")
            loComandoSeleccionar.AppendLine("           AND Fec_Ini         BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Cli         BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Ven         BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Status      IN ('Afectado', 'Pagado')")
            loComandoSeleccionar.AppendLine("           AND Cod_Mon         BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Rev         BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cod_Suc         BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY  DATEPART(YEAR, Fec_Ini),DATEPART(MONTH, Fec_Ini), DATEPART(DAY, Fec_Ini) ")
            
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(YEAR, Cobros.fec_ini) AS   Año,")
            loComandoSeleccionar.AppendLine("           DATEPART(MONTH, Cobros.fec_ini)AS Mes,")
            loComandoSeleccionar.AppendLine("           DATEPART(DAY, Cobros.fec_ini)AS Dia,")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Efectivo'       THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Efectivo, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Ticket'         THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Ticket, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Cheque'         THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Cheque, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Tarjeta'        THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Tarjeta, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Deposito'       THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Deposito, ")
            loComandoSeleccionar.AppendLine("           SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Transferencia'  THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Transferencia ")
            loComandoSeleccionar.AppendLine(" INTO      #Temporal2 ")
            loComandoSeleccionar.AppendLine(" FROM      Cobros Cobros")
            loComandoSeleccionar.AppendLine("           JOIN Vendedores AS Vendedores ON  Vendedores.Cod_Ven = Cobros.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           JOIN Detalles_Cobros AS Detalles_Cobros ON Detalles_Cobros.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine(" WHERE     Cobros.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Cli  BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Ven  BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Status   IN ('Confirmado')")
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Mon  BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Rev  BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Suc  BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY  DATEPART(YEAR, Cobros.Fec_Ini),DATEPART(MONTH, Cobros.Fec_Ini), DATEPART(DAY, Cobros.Fec_Ini)")
            
            loComandoSeleccionar.AppendLine(" SELECT    ISNULL(#temporal.Año, #temporal2.Año)   AS Año, ")
            loComandoSeleccionar.AppendLine("           ISNULL(#temporal.Mes, #temporal2.Mes)   AS Mes, ")
            loComandoSeleccionar.AppendLine("           ISNULL(#temporal.Dia, #temporal2.Dia)   AS Dia, ")
            loComandoSeleccionar.AppendLine("           ISNULL(#temporal.Mon_Net,0)		        AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("           (ISNULL(#temporal2.Efectivo,0))      AS Efectivo, ")
            loComandoSeleccionar.AppendLine("           (ISNULL(#temporal2.Ticket,0))        AS Ticket, ")
            loComandoSeleccionar.AppendLine("           (ISNULL(#temporal2.Cheque,0))        AS Cheque, ")
            loComandoSeleccionar.AppendLine("           (ISNULL(#temporal2.Tarjeta,0))       AS Tarjeta, ")
            loComandoSeleccionar.AppendLine("           (ISNULL(#temporal2.Deposito,0))      AS Deposito, ")
            loComandoSeleccionar.AppendLine("           (ISNULL(#temporal2.Transferencia,0)) AS Transferencia,")
            loComandoSeleccionar.AppendLine("           (ISNULL(#temporal2.Efectivo,0) + ISNULL(#temporal2.Cheque,0) + ISNULL(#temporal2.Tarjeta,0) + ISNULL(#temporal2.Deposito,0) + ISNULL(#temporal2.Transferencia,0) + ISNULL(#temporal2.Ticket,0)) AS Total_Cobros")
            loComandoSeleccionar.AppendLine(" INTO      #Temporal002 ")
            loComandoSeleccionar.AppendLine(" FROM      #Temporal #Temporal   ")
            loComandoSeleccionar.AppendLine("           FULL JOIN  #temporal2 AS #temporal2 ON ((#temporal.Año = #temporal2.Año) AND (#temporal.Mes = #temporal2.Mes) AND (#temporal.Dia = #temporal2.Dia)) ")
            
            
            loComandoSeleccionar.AppendLine(" SELECT    'Cuantos'  AS  Cuantos, ")
            loComandoSeleccionar.AppendLine("           Año, ")
            loComandoSeleccionar.AppendLine("           Mes, ")
            loComandoSeleccionar.AppendLine("           Dia, ")
            loComandoSeleccionar.AppendLine("           Mon_Net, ")
            loComandoSeleccionar.AppendLine("           Efectivo, ")
            loComandoSeleccionar.AppendLine("           Ticket, ")
            loComandoSeleccionar.AppendLine("           Cheque, ")
            loComandoSeleccionar.AppendLine("           Tarjeta, ")
            loComandoSeleccionar.AppendLine("           Deposito, ")
            loComandoSeleccionar.AppendLine("           Transferencia,")
            loComandoSeleccionar.AppendLine("           Total_Cobros")
            loComandoSeleccionar.AppendLine(" FROM      #Temporal002 ")
            loComandoSeleccionar.AppendLine(" ORDER BY  " & lcOrdenamiento)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTCobros_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTCobros_Fechas.ReportSource = loObjetoReporte

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
' CMS: 08/07/09: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' CMS: 20/07/09: Se agregaron las columnas ticket y transferencia							'
'-------------------------------------------------------------------------------------------'
' RJG: 28/06/11: Corrección de agrupación de monto neto.									'	
'-------------------------------------------------------------------------------------------'
