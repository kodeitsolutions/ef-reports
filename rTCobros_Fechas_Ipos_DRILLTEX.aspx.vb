'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTCobros_Fechas_Ipos_DRILLTEX"
'-------------------------------------------------------------------------------------------'
Partial Class rTCobros_Fechas_Ipos_DRILLTEX
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
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()
            
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  CAST(CONVERT(VARCHAR(30), Cuentas_Cobrar.Fec_Ini, 102) AS DATE) AS Fecha,")
            loConsulta.AppendLine("        SUM(CASE WHEN Cuentas_Cobrar.Cod_Tip IN ('FACT', 'ATD')")
            'loConsulta.AppendLine("            AND Cuentas_Cobrar.Ipos = '1' ")
            loConsulta.AppendLine("            THEN Cuentas_Cobrar.mon_bru - mon_des + mon_rec")
            loConsulta.AppendLine("            ELSE 0 ")
            loConsulta.AppendLine("        END)                        AS Facturas_Bruto, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Cuentas_Cobrar.Cod_Tip IN ('FACT', 'ATD')")
            'loConsulta.AppendLine("            AND Cuentas_Cobrar.Ipos = '1' ")
            loConsulta.AppendLine("            THEN Cuentas_Cobrar.mon_net ")
            loConsulta.AppendLine("            ELSE 0 ")
            loConsulta.AppendLine("        END)                        AS Facturas_Neto, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Cuentas_Cobrar.Cod_Tip IN ('N/CR')")
            'loConsulta.AppendLine("            AND Cuentas_Cobrar.Ipos = '1' ")
            loConsulta.AppendLine("            THEN Cuentas_Cobrar.mon_net ")
            loConsulta.AppendLine("            ELSE 0 ")
            loConsulta.AppendLine("        END)                        AS Notas_Credito, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Cuentas_Cobrar.Cod_Tip IN ('RETIVA')")
            loConsulta.AppendLine("            THEN Cuentas_Cobrar.mon_net ")
            loConsulta.AppendLine("            ELSE 0 ")
            loConsulta.AppendLine("        END)                        AS Retenciones ")
            loConsulta.AppendLine("INTO    #tmpCxC")
            loConsulta.AppendLine("FROM    Cuentas_Cobrar")
            loConsulta.AppendLine("WHERE   Cuentas_Cobrar.Cod_Tip     IN ('Fact','ATD')")
            'loConsulta.AppendLine("    AND Cuentas_Cobrar.Ipos = '1'")
            loConsulta.AppendLine("    AND Cuentas_Cobrar.Fec_Ini         BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("    AND " & lcParametro0Hasta)
            loConsulta.AppendLine("    AND Cod_Cli         BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("    AND " & lcParametro1Hasta)
            loConsulta.AppendLine("    AND Cod_Ven         BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("    AND " & lcParametro2Hasta)
            loConsulta.AppendLine("    AND Cuentas_Cobrar.Status      IN ('Afectado', 'Pagado')")
            loConsulta.AppendLine("    AND Cod_Mon         BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("    AND " & lcParametro3Hasta)
            loConsulta.AppendLine("    AND Cod_Rev         BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("    AND " & lcParametro4Hasta)
            loConsulta.AppendLine("    AND Cod_Suc         BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("    AND " & lcParametro5Hasta)
            loConsulta.AppendLine("    AND Cuentas_Cobrar.Usu_Cre     BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("    AND " & lcParametro6Hasta)
            loConsulta.AppendLine("GROUP BY CAST(CONVERT(VARCHAR(30), Cuentas_Cobrar.Fec_Ini, 102) AS DATE)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  CAST(CONVERT(VARCHAR(30), Cobros.fec_ini, 102) AS DATE) AS Fecha,")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Efectivo'       THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Efectivo, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Ticket'         THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Ticket, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Cheque'         THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Cheque, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Tarjeta'        THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Tarjeta, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Deposito'       THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Deposito, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Transferencia'  THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Transferencia ")
            loConsulta.AppendLine("INTO    #tmpCobros ")
            loConsulta.AppendLine("FROM    Cobros Cobros")
            loConsulta.AppendLine("    JOIN Vendedores ON Vendedores.Cod_Ven = Cobros.Cod_Ven ")
            loConsulta.AppendLine("    JOIN Detalles_Cobros ON Detalles_Cobros.Documento = Cobros.Documento")
            loConsulta.AppendLine("WHERE    Cobros.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("     AND " & lcParametro0Hasta)
            'loConsulta.AppendLine("	    AND Cobros.Ipos = '1'")
            loConsulta.AppendLine("     AND Cobros.Cod_Cli  BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("     AND " & lcParametro1Hasta)
            loConsulta.AppendLine("     AND Cobros.Cod_Ven  BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("     AND " & lcParametro2Hasta)
            loConsulta.AppendLine("     AND Cobros.Status   IN ('Confirmado')")
            loConsulta.AppendLine("     AND Cobros.Cod_Mon  BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("     AND " & lcParametro3Hasta)
            loConsulta.AppendLine("     AND Cobros.Cod_Rev  BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("     AND " & lcParametro4Hasta)
            loConsulta.AppendLine("     AND Cobros.Cod_Suc  BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("     AND " & lcParametro5Hasta)
            loConsulta.AppendLine("     AND Cobros.Cod_Usu     BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("     AND " & lcParametro6Hasta)
            loConsulta.AppendLine("GROUP BY  CAST(CONVERT(VARCHAR(30), Cobros.Fec_Ini, 102) AS DATE)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  COALESCE(#tmpCxC.Fecha, #tmpCobros.Fecha)       AS Fecha, ")
            loConsulta.AppendLine("        COALESCE(#tmpCxC.Facturas_Bruto, 0)		        AS Facturas_Bruto, ")
            loConsulta.AppendLine("        COALESCE(#tmpCxC.Facturas_Neto, 0)			    AS Facturas_Neto, ")
            loConsulta.AppendLine("        COALESCE(#tmpCxC.Notas_Credito, 0)			    AS Notas_Credito, ")
            loConsulta.AppendLine("        COALESCE(#tmpCxC.Retenciones, 0)			    AS Retenciones, ")
            loConsulta.AppendLine("        COALESCE(#tmpCobros.Efectivo, 0)                AS Efectivo, ")
            loConsulta.AppendLine("        COALESCE(#tmpCobros.Ticket, 0)                  AS Ticket, ")
            loConsulta.AppendLine("        COALESCE(#tmpCobros.Cheque, 0)                  AS Cheque, ")
            loConsulta.AppendLine("        COALESCE(#tmpCobros.Tarjeta, 0)                 AS Tarjeta, ")
            loConsulta.AppendLine("        COALESCE(#tmpCobros.Deposito, 0)                AS Deposito, ")
            loConsulta.AppendLine("        COALESCE(#tmpCobros.Transferencia, 0)           AS Transferencia,")
            loConsulta.AppendLine("        COALESCE(#tmpCobros.Efectivo, 0) ")
            loConsulta.AppendLine("        + COALESCE(#tmpCobros.Cheque, 0) ")
            loConsulta.AppendLine("        + COALESCE(#tmpCobros.Tarjeta, 0) ")
            loConsulta.AppendLine("        + COALESCE(#tmpCobros.Deposito, 0) ")
            loConsulta.AppendLine("        + COALESCE(#tmpCobros.Transferencia, 0) ")
            loConsulta.AppendLine("        + COALESCE(#tmpCobros.Ticket, 0)                AS Total_Cobros")
            loConsulta.AppendLine("FROM    #tmpCxC   ")
            loConsulta.AppendLine("FULL JOIN  #tmpCobros ON (#tmpCxC.Fecha = #tmpCobros.Fecha)  ")
            loConsulta.AppendLine("ORDER BY  Fecha ASC")
            'loConsulta.AppendLine("ORDER BY  " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- DROP TABLE #tmpCxC;")
            loConsulta.AppendLine("-- DROP TABLE #tmpCobros;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTCobros_Fechas_Ipos_DRILLTEX", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTCobros_Fechas_Ipos_DRILLTEX.ReportSource = loObjetoReporte

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
' RJG: 31/10/14: Programacion inicial, a partir de rTCobros_Fechas_Ipos.                    '
'-------------------------------------------------------------------------------------------'
