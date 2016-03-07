'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTCobros_Clientes_iPos"
'-------------------------------------------------------------------------------------------'
Partial Class rTCobros_Clientes_iPos
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
            loConsulta.AppendLine("CREATE TABLE #tmpVentas(Cod_Cli CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                        Mon_Net DECIMAL(18, 10));")
            loConsulta.AppendLine("CREATE TABLE #tmpCobros(Cod_Cli CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                        Efectivo DECIMAL(18, 10), ")
            loConsulta.AppendLine("                        Ticket DECIMAL(18, 10), ")
            loConsulta.AppendLine("                        Cheque DECIMAL(18, 10), ")
            loConsulta.AppendLine("                        Tarjeta DECIMAL(18, 10), ")
            loConsulta.AppendLine("                        Deposito DECIMAL(18, 10), ")
            loConsulta.AppendLine("                        Transferencia DECIMAL(18, 10));")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO  #tmpVentas(Cod_Cli, Mon_Net)")
            loConsulta.AppendLine("SELECT       Cuentas_Cobrar.Cod_Cli          AS Cod_Cli,")
            loConsulta.AppendLine("             SUM(Cuentas_Cobrar.Mon_Net)     AS Mon_Net ")
            loConsulta.AppendLine("FROM         Cuentas_Cobrar ")
            loConsulta.AppendLine("WHERE        Cuentas_Cobrar.Cod_Tip IN ('FACT','ATD')")
            loConsulta.AppendLine("			    AND Cuentas_Cobrar.Ipos = '1'")
            loConsulta.AppendLine("             AND Cuentas_Cobrar.Fec_Ini  BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("             AND " & lcParametro0Hasta)
            loConsulta.AppendLine("             AND Cuentas_Cobrar.Cod_Cli  BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("             AND " & lcParametro1Hasta)
            loConsulta.AppendLine("             AND Cuentas_Cobrar.Cod_Ven  BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("             AND " & lcParametro2Hasta)
            loConsulta.AppendLine("             AND Cuentas_Cobrar.Status      IN ('Afectado', 'Pagado')")
            loConsulta.AppendLine("             AND Cuentas_Cobrar.Cod_Mon  BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("             AND " & lcParametro3Hasta)
            loConsulta.AppendLine("             AND Cuentas_Cobrar.Cod_Rev  BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("             AND " & lcParametro4Hasta)
            loConsulta.AppendLine("             AND Cuentas_Cobrar.Cod_Suc  BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("             AND " & lcParametro5Hasta)
            loConsulta.AppendLine("             AND Cuentas_Cobrar.Usu_Cre  BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("             AND " & lcParametro6Hasta)
            loConsulta.AppendLine(" GROUP BY  Cuentas_Cobrar.Cod_Cli ")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpCobros(Cod_Cli, Efectivo, Ticket, Cheque, Tarjeta, Deposito, Transferencia)")
            loConsulta.AppendLine("SELECT  Cobros.Cod_Cli                                                                                    AS Cod_Cli, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Efectivo'       THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Efectivo, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Ticket'         THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Ticket, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Cheque'         THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Cheque, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Tarjeta'        THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Tarjeta, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Deposito'       THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Deposito, ")
            loConsulta.AppendLine("        SUM(CASE WHEN Detalles_Cobros.Tip_Ope = 'Transferencia'  THEN Detalles_Cobros.Mon_Net ELSE 0 END) AS Transferencia ")
            loConsulta.AppendLine("FROM    Cobros ")
            loConsulta.AppendLine("    JOIN Detalles_Cobros ON Detalles_Cobros.Documento = Cobros.Documento")
            loConsulta.AppendLine("WHERE    Cobros.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("         AND " & lcParametro0Hasta)
            loConsulta.AppendLine("			AND Cobros.Ipos = '1'")
            loConsulta.AppendLine("         AND Cobros.Cod_Cli  BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("         AND " & lcParametro1Hasta)
            loConsulta.AppendLine("         AND Cobros.Cod_Ven  BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("         AND " & lcParametro2Hasta)
            loConsulta.AppendLine("         AND Cobros.Status   IN ('Confirmado')")
            loConsulta.AppendLine("         AND Cobros.Cod_Mon  BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("         AND " & lcParametro3Hasta)
            loConsulta.AppendLine("         AND Cobros.Cod_Rev  BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("         AND " & lcParametro4Hasta)
            loConsulta.AppendLine("         AND Cobros.Cod_Suc  BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("         AND " & lcParametro5Hasta)
            loConsulta.AppendLine("         AND Cobros.Usu_Cre  BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("         AND " & lcParametro6Hasta)
            loConsulta.AppendLine(" GROUP BY  Cobros.Cod_Cli ")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Clientes.Cod_Cli                        AS Cod_Cli, ")
            loConsulta.AppendLine("            Clientes.Nom_Cli                        AS Nom_Cli, ")
            loConsulta.AppendLine("            #tmpVentas.Mon_Net                      AS Mon_Net, ") 
            loConsulta.AppendLine("            #tmpCobros.Efectivo                     AS Efectivo, ")
            loConsulta.AppendLine("            #tmpCobros.Ticket                       AS Ticket, ")
            loConsulta.AppendLine("            #tmpCobros.Cheque                       AS Cheque, ")
            loConsulta.AppendLine("            #tmpCobros.Tarjeta                      AS Tarjeta, ")
            loConsulta.AppendLine("            #tmpCobros.Deposito                     AS Deposito, ")
            loConsulta.AppendLine("            #tmpCobros.Transferencia                AS Transferencia, ")
            loConsulta.AppendLine("            ( COALESCE(#tmpCobros.Efectivo, 0) ")
            loConsulta.AppendLine("            + COALESCE(#tmpCobros.Cheque, 0) ")
            loConsulta.AppendLine("            + COALESCE(#tmpCobros.Tarjeta,0) ")
            loConsulta.AppendLine("            + COALESCE(#tmpCobros.Deposito,0) ")
            loConsulta.AppendLine("            + COALESCE(#tmpCobros.Transferencia, 0) ")
            loConsulta.AppendLine("            + COALESCE(#tmpCobros.Ticket,0))        AS Total_Cobros")
            loConsulta.AppendLine("FROM        #tmpVentas") 
            loConsulta.AppendLine("FULL JOIN   #tmpCobros ON (#tmpCobros.Cod_Cli = #tmpVentas.Cod_Cli) ")
            loConsulta.AppendLine("     JOIN   Clientes ON Clientes.cod_cli = #tmpVentas.Cod_Cli")
            loConsulta.AppendLine("ORDER BY    Clientes." & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

           

			'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTCobros_Clientes_iPos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTCobros_Clientes_iPos.ReportSource = loObjetoReporte

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
' MAT: 21/07/11: Ajuste del Select															' 
'-------------------------------------------------------------------------------------------'
