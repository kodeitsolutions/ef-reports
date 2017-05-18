'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rLEstadoCuenta_HistoricoVentas"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rLEstadoCuenta_HistoricoVentas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            If (lcOrdenamiento.ToLower().Trim() = "cod_cli") Then
                lcOrdenamiento = "#tmpDocumentos.Cod_Cli ASC"
            End If
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE	@ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE	@ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loConsulta.AppendLine("DECLARE	@lcCodCli_Desde	AS VARCHAR(10) = " & lcParametro1Desde)
            loConsulta.AppendLine("DECLARE	@lcCodCli_Hasta	AS VARCHAR(10) = " & lcParametro1Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpClientes(  Cod_Cli CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Nom_Cli CHAR(100) COLLATE DATABASE_DEFAULT);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE INDEX PK_tmpClientes_Cod_Cli ON #tmpClientes(Cod_Cli) INCLUDE (Nom_Cli);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpClientes(Cod_Cli, Nom_Cli)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT Cod_Cli, Nom_Cli ")
            loConsulta.AppendLine("FROM Clientes")
            loConsulta.AppendLine("WHERE Cod_Cli BETWEEN @lcCodCli_Desde AND @lcCodCli_Hasta")
            loConsulta.AppendLine("        ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpDocumentos(Documento    CHAR(10)    COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Cod_Cli      CHAR(10)    COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Cod_Ven      CHAR(10)    COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Fecha        DATE, ")
            loConsulta.AppendLine("                            Referencia   VARCHAR(MAX), ")
            loConsulta.AppendLine("                            Debe         DECIMAL(28,10), ")
            loConsulta.AppendLine("                            Haber        DECIMAL(28,10), ")
            loConsulta.AppendLine("                            Mon_Sal      Decimal(28,10), ")
            loConsulta.AppendLine("                            Tipo         CHAR(10)   COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Cod_Tip      CHAR(10)   COLLATE DATABASE_DEFAULT);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpDocumentos(Documento, Cod_Cli,  Fecha, Referencia, Debe, Haber, Mon_Sal, Tipo)")
            loConsulta.AppendLine("SELECT   Cuentas_Cobrar.Documento                            AS Documento, ")
            loConsulta.AppendLine("         Cuentas_Cobrar.Cod_Cli                              AS Cod_Cli, ")
            loConsulta.AppendLine("         Cuentas_Cobrar.Fec_Ini                              AS Fecha,")
            loConsulta.AppendLine("         CASE WHEN Cuentas_Cobrar.Cod_Tip = 'ADEL'")
            loConsulta.AppendLine("         	 THEN (SELECT CONCAT(RTRIM(Tip_Ope),' ', RTRIM(Num_Doc))")
            loConsulta.AppendLine("         		   FROM Detalles_Cobros")
            loConsulta.AppendLine("         		   WHERE Documento = Cuentas_Cobrar.Doc_Ori) ")
            loConsulta.AppendLine("         	 ELSE Cuentas_Cobrar.Comentario")
            loConsulta.AppendLine("         END													AS Referencia,")
            loConsulta.AppendLine("         CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' ")
            loConsulta.AppendLine("              THEN Cuentas_Cobrar.Mon_Net")
            loConsulta.AppendLine("              ELSE 0")
            loConsulta.AppendLine("         END                                                 AS Debe,")
            loConsulta.AppendLine("         CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito'")
            loConsulta.AppendLine("              THEN Cuentas_Cobrar.Mon_Net")
            loConsulta.AppendLine("              ELSE 0")
            loConsulta.AppendLine("         END                                                 AS Haber,")
            loConsulta.AppendLine("         CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' ")
            loConsulta.AppendLine("              THEN Cuentas_Cobrar.Mon_Sal")
            loConsulta.AppendLine("              ELSE -Cuentas_Cobrar.Mon_Sal")
            loConsulta.AppendLine("         END                                                 AS Mon_Sal,")
            loConsulta.AppendLine("         Cuentas_Cobrar.Cod_Tip                              AS Tipo")
            loConsulta.AppendLine("FROM Cuentas_Cobrar")
            loConsulta.AppendLine(" JOIN #tmpClientes ON  #tmpClientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loConsulta.AppendLine("     AND Cuentas_Cobrar.Status <> 'Anulado'")
            loConsulta.AppendLine("     AND Fec_Ini <= @ldFecha_Hasta")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UNION ALL ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   Cobros.Documento                                    AS Documento, ")
            loConsulta.AppendLine("         Cobros.Cod_Cli                                      AS Cod_Cli, ")
            loConsulta.AppendLine("         Cobros.Fec_Ini                                      AS Fecha, ")
            loConsulta.AppendLine("         Cobros.Comentario                                   AS Referencia,")
            loConsulta.AppendLine("         -Cobros.Mon_Net                                     AS Debe,")
            loConsulta.AppendLine("         0                                                   AS Haber,")
            loConsulta.AppendLine("         0                                                   AS Mon_Sal,")
            loConsulta.AppendLine("         'COBRO'                                             AS Tipo")
            loConsulta.AppendLine("FROM        Cobros")
            loConsulta.AppendLine("    JOIN    #tmpClientes ON #tmpClientes.Cod_Cli = Cobros.Cod_Cli")
            loConsulta.AppendLine("WHERE       Cobros.Automatico = 0")
            loConsulta.AppendLine("		AND Cobros.Status IN ('Confirmado')")
            loConsulta.AppendLine("    AND     Fec_Ini <= @ldFecha_Hasta")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--PENDIENTES POR CONCILIAR")
            loConsulta.AppendLine("SELECT COUNT(Documento)								AS Pendientes,")
            loConsulta.AppendLine("		Cod_Cli											AS Cod_Cli")
            loConsulta.AppendLine("INTO #tmpMovPendientesFact")
            loConsulta.AppendLine("FROM Cuentas_Cobrar")
            loConsulta.AppendLine("WHERE Cod_Tip = 'FACT' AND Status <> 'Pagado' AND Mon_Sal > 0 ")
            loConsulta.AppendLine("	AND Cod_Cli BETWEEN @lcCodCli_Desde AND @lcCodCli_Hasta")
            loConsulta.AppendLine("	AND Fec_Ini < @ldFecha_Hasta")
            loConsulta.AppendLine("GROUP BY Cod_Cli")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT COUNT(Documento)								AS Pendientes,")
            loConsulta.AppendLine("		Cod_Cli											AS Cod_Cli")
            loConsulta.AppendLine("INTO #tmpMovPendientesAdel")
            loConsulta.AppendLine("FROM Cuentas_Cobrar")
            loConsulta.AppendLine("WHERE Cod_Tip = 'ADEL' AND Status <> 'Pagado' AND Mon_Sal > 0 ")
            loConsulta.AppendLine("	AND Cod_Cli BETWEEN @lcCodCli_Desde AND @lcCodCli_Hasta")
            loConsulta.AppendLine("	AND Fec_Ini < @ldFecha_Hasta")
            loConsulta.AppendLine("GROUP BY Cod_Cli")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   #tmpDocumentos.Documento                                            AS Documento, ")
            loConsulta.AppendLine("         #tmpDocumentos.Cod_Cli                                              AS Cod_Cli, ")
            loConsulta.AppendLine("         #tmpClientes.Nom_Cli                                                AS Nom_Cli, ")
            loConsulta.AppendLine("         #tmpDocumentos.Fecha                                                AS Fecha,")
            loConsulta.AppendLine("         CASE WHEN LEN(#tmpDocumentos.Referencia) > 60")
            loConsulta.AppendLine("              THEN CONCAT(SUBSTRING(#tmpDocumentos.Referencia,1,60),'...')")
            loConsulta.AppendLine("              ELSE #tmpDocumentos.Referencia")
            loConsulta.AppendLine("         END                                                                 AS Referencia,")
            loConsulta.AppendLine("         #tmpDocumentos.Debe                                                 AS Mon_Deb,")
            loConsulta.AppendLine("         #tmpDocumentos.Haber                                                AS Mon_Hab,")
            loConsulta.AppendLine("         #tmpDocumentos.Mon_Sal                                              AS Mon_Sal,")
            loConsulta.AppendLine("         #tmpDocumentos.Tipo                                                 AS Tipo,")
            loConsulta.AppendLine("         COALESCE(Iniciales.Monto, 0)                                        AS Mon_Ini,")
            loConsulta.AppendLine("         COALESCE(Iniciales.Monto, 0) ")
            loConsulta.AppendLine("         + SUM(#tmpDocumentos.Debe - #tmpDocumentos.Haber) ")
            loConsulta.AppendLine("              OVER( PARTITION BY #tmpDocumentos.Cod_Cli")
            loConsulta.AppendLine("                    ORDER BY #tmpDocumentos.Fecha, #tmpDocumentos.Cod_Tip")
            loConsulta.AppendLine("                    ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW ")
            loConsulta.AppendLine("              )                                                              AS Saldo_Acumulado, ")
            loConsulta.AppendLine("         COALESCE(Iniciales.Monto, 0) ")
            loConsulta.AppendLine("         + SUM(#tmpDocumentos.Debe - #tmpDocumentos.Haber) ")
            loConsulta.AppendLine("              OVER( PARTITION BY #tmpDocumentos.Cod_Cli)                     AS Saldo_Cliente, ")
            loConsulta.AppendLine("         COALESCE(Iniciales.Monto, 0) ")
            loConsulta.AppendLine("         + SUM(#tmpDocumentos.Debe - #tmpDocumentos.Haber) OVER()            AS Saldo_Total,")
            loConsulta.AppendLine("		    COALESCE((#tmpMovPendientesFact.Pendientes),0)						AS PendientesFact,")
            loConsulta.AppendLine("		    COALESCE((#tmpMovPendientesAdel.Pendientes),0)						AS PendientesAdel,")
            loConsulta.AppendLine("         @ldFecha_Desde                                                      AS Desde,")
            loConsulta.AppendLine("         @ldFecha_Hasta                                                      AS Hasta")
            loConsulta.AppendLine("FROM #tmpDocumentos")
            loConsulta.AppendLine(" LEFT JOIN (SELECT Cod_Cli, SUM(#tmpDocumentos.Debe - #tmpDocumentos.Haber)  AS Monto ")
            loConsulta.AppendLine("            FROM    #tmpDocumentos")
            loConsulta.AppendLine("            WHERE   Fecha < @ldFecha_Desde")
            loConsulta.AppendLine("            GROUP BY Cod_Cli")
            loConsulta.AppendLine("            ) Iniciales")
            loConsulta.AppendLine("    ON #tmpDocumentos.Cod_Cli = Iniciales.Cod_Cli")
            loConsulta.AppendLine("	LEFT JOIN #tmpMovPendientesFact ON #tmpMovPendientesFact.Cod_Cli = #tmpDocumentos.Cod_Cli")
            loConsulta.AppendLine("	LEFT JOIN #tmpMovPendientesAdel ON #tmpMovPendientesAdel.Cod_Cli = #tmpDocumentos.Cod_Cli")
            loConsulta.AppendLine(" JOIN #tmpClientes ON #tmpClientes.Cod_Cli = #tmpDocumentos.Cod_Cli")
            loConsulta.AppendLine("WHERE #tmpDocumentos.Fecha >= @ldFecha_Desde")
            loConsulta.AppendLine("ORDER BY #tmpDocumentos.Cod_Cli, #tmpDocumentos.Fecha")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DROP TABLE #tmpClientes")
            loConsulta.AppendLine("DROP TABLE #tmpDocumentos")
            loConsulta.AppendLine("DROP TABLE #tmpMovPendientesAdel")
            loConsulta.AppendLine("DROP TABLE #tmpMovPendientesFact")
            loConsulta.AppendLine("")


            'loComandoSeleccionar.AppendLine("EXEC   sp_Estado_de_Cuenta_Ventas")
            'loComandoSeleccionar.AppendLine("       @sp_FecIni          = " & lcParametro0Desde & ",")
            'loComandoSeleccionar.AppendLine("       @sp_FecFin          = " & lcParametro0Hasta & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodCli_Desde    = " & lcParametro1Desde & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodCli_Hasta    = " & lcParametro1Hasta & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodCla_Desde    = " & lcParametro2Desde & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodCla_Hasta    = " & lcParametro2Hasta & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodTip_Desde    = " & lcParametro3Desde & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodTip_Hasta    = " & lcParametro3Hasta & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodZon_Desde    = " & lcParametro5Desde & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodZon_Hasta    = " & lcParametro5Hasta & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodVen_Desde    = " & lcParametro4Desde & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodVen_Hasta    = " & lcParametro4Hasta & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodMon_Desde    = " & lcParametro6Desde & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodMon_Hasta    = " & lcParametro6Hasta & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodSuc_Desde    = " & lcParametro7Desde & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodSuc_Hasta    = " & lcParametro7Hasta & ",")
            'loComandoSeleccionar.AppendLine("       @sp_TipRev          = " & lcParametro8Desde & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodRev_Desde    = " & lcParametro9Desde & ",")
            'loComandoSeleccionar.AppendLine("       @sp_CodRev_Hasta    = " & lcParametro9Hasta & ",")
            'loComandoSeleccionar.AppendLine("       @sp_Ordenamiento    = '" & lcOrdenamiento & "'")

            'Me.mEscribirConsulta(loConsulta.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rLEstadoCuenta_HistoricoVentas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvTRF_rLEstadoCuenta_HistoricoVentas.ReportSource = loObjetoReporte

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
' DLC: 21/07/2010: Programacion inicial
'-------------------------------------------------------------------------------------------'
' DLC: 02/09/2010: Cambio de la consulta a procedimiento almacenado.
'                   - Se ajusto los filtro de las consultas.
'                   - Se ajusto la forma en que se calcula los cobros.
'                   - Se ajusto la forma en que se calcula las cuentas por cobrar.
'-------------------------------------------------------------------------------------------'
' DLC: 15/09/2010: Ajuste en la forma de obtener los detalles de Cobros, asi como también,
'                ajustar en el RPT, la forma de mostrar los detalles de Cobros.
'-------------------------------------------------------------------------------------------'
' MAT: 13/04/11: Reprogramación del Reporte y su respectivo Store Procedure
'-------------------------------------------------------------------------------------------'
' MAT: 14/04/11: Ajuste de la vista de Diseño.
'-------------------------------------------------------------------------------------------'
' MAT: 27/04/11: Se elimino el filtro Detalle
'-------------------------------------------------------------------------------------------'
' RJG: 29/01/16: se movió la consulta del SP al reporte, y se ajustaron las uniones y filtros
'-------------------------------------------------------------------------------------------'
' EAG: 25/02/16: Se acomodaron los totales
'-------------------------------------------------------------------------------------------'
