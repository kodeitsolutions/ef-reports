'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rAnalitico_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rAnalitico_Clientes
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecIni			AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecFin			AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE	@sp_CodCli_Desde	AS VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE	@sp_CodCli_Hasta	AS VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero	AS DECIMAL(28, 10) = 0;")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio	AS VARCHAR(10) = '';")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--SALDO INICIAL")
            loComandoSeleccionar.AppendLine("SELECT	Cuentas_Cobrar.Cod_Cli							AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Cobrar.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE -Cuentas_Cobrar.Mon_Net	")
            loComandoSeleccionar.AppendLine("			END)										AS Sal_Ini")
            loComandoSeleccionar.AppendLine("INTO #tmpSaldoInicial")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("	JOIN Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Cobrar.Fec_Ini < @sp_FecIni")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Cobrar.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("GROUP BY Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Cobros.Cod_Cli                                  AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Cobros.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN Renglones_Cobros.Mon_Abo")
            loComandoSeleccionar.AppendLine("				ELSE -Renglones_Cobros.Mon_Abo	")
            loComandoSeleccionar.AppendLine("			END) -(Cobros.Mon_Ret - Cobros.Mon_Des)     AS Sal_Ini")
            loComandoSeleccionar.AppendLine("FROM Cobros")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Cobros ON Cobros.Documento = Renglones_Cobros.Documento		")
            loComandoSeleccionar.AppendLine("	JOIN Clientes ON Clientes.Cod_Cli = Cobros.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE  Cobros.Status IN ('Confirmado')")
            loComandoSeleccionar.AppendLine("	AND Cobros.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("	AND	Cobros.Fec_Ini < @sp_FecIni")
            loComandoSeleccionar.AppendLine("	AND Cobros.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("GROUP BY Cobros.Cod_Cli, Cobros.Mon_Ret, Cobros.Mon_Des")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--MOVIMIENTOS")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Cobros'										AS Tabla,")
            loComandoSeleccionar.AppendLine("		0												AS Orden,")
            loComandoSeleccionar.AppendLine("		Clientes.Cod_Cli								AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli								AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		'COBRO'											AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cobros.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("		Cobros.Fec_Ini									AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Cobros.Comentario								AS Referencia,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Cobros.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("				THEN	Renglones_Cobros.Mon_Abo	")
            loComandoSeleccionar.AppendLine("				ELSE	@lnCero")
            loComandoSeleccionar.AppendLine("			END)+(Cobros.Mon_Ret + Cobros.Mon_Des)  	AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Cobros.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("				THEN	@lnCero	")
            loComandoSeleccionar.AppendLine("				ELSE	Renglones_Cobros.Mon_Abo	")
            loComandoSeleccionar.AppendLine("			END)  		AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Mon_Sal")
            loComandoSeleccionar.AppendLine("INTO #tmpMovimientos")
            loComandoSeleccionar.AppendLine("FROM	Cobros")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Cobros ON Cobros.Documento = Renglones_Cobros.Documento")
            loComandoSeleccionar.AppendLine("JOIN	Clientes ON Clientes.Cod_Cli = Cobros.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE	Cobros.Status IN ('Confirmado')")
            loComandoSeleccionar.AppendLine("		AND	Cobros.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("		AND Cobros.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("		AND Cobros.Automatico = 0")
            loComandoSeleccionar.AppendLine("GROUP BY	Clientes.Cod_Cli, Clientes.Nom_Cli, Cobros.Cod_Ven, Cobros.Documento,")
            loComandoSeleccionar.AppendLine("			Cobros.Fec_Ini, Cobros.Registro, Cobros.Comentario, Cobros.Mon_Ret, Cobros.Mon_Des")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Cuentas_Cobrar'								AS Tabla,")
            loComandoSeleccionar.AppendLine("		0												AS Orden,")
            loComandoSeleccionar.AppendLine("		Clientes.Cod_Cli								AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli								AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Tip							AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Documento						AS Documento,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Ini							AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		CASE WHEN Cuentas_Cobrar.Cod_Tip = 'ADEL' THEN")
            loComandoSeleccionar.AppendLine("            (SELECT CONCAT(RTRIM(Tip_Ope),' ', RTRIM(Num_Doc))")
            loComandoSeleccionar.AppendLine("            FROM Detalles_Cobros")
            loComandoSeleccionar.AppendLine("            WHERE Documento = Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("            ) ")
            loComandoSeleccionar.AppendLine("            ELSE Cuentas_Cobrar.Comentario")
            loComandoSeleccionar.AppendLine("        END											AS Referencia,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Cobrar.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN @lnCero")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Cobrar.Mon_Net")
            loComandoSeleccionar.AppendLine("			END)										AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Mon_Sal")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes  ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE	Cuentas_Cobrar.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Tip <> 'FACT'")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Cuentas_Cobrar'								AS Tabla,")
            loComandoSeleccionar.AppendLine("		0												AS Orden,")
            loComandoSeleccionar.AppendLine("		Clientes.Cod_Cli								AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli								AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Tip							AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Documento						AS Documento,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Ini							AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Comentario						AS Referencia,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN (CASE  WHEN Cuentas_Cobrar.Mon_Sal > 0")
            loComandoSeleccionar.AppendLine("							THEN Cuentas_Cobrar.Mon_Sal")
            loComandoSeleccionar.AppendLine("							ELSE Cuentas_Cobrar.Mon_Net")
            loComandoSeleccionar.AppendLine("					  END)")
            loComandoSeleccionar.AppendLine("				ELSE @lnCero")
            loComandoSeleccionar.AppendLine("		END)											AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN @lnCero")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Cobrar.Mon_Net")
            loComandoSeleccionar.AppendLine("		END)											AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Mon_Sal")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes  ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE	Cuentas_Cobrar.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--PENDIENTES POR CONCILIAR")
            loComandoSeleccionar.AppendLine("SELECT COUNT(Documento)								AS Pendientes,")
            loComandoSeleccionar.AppendLine("		Cod_Cli											AS Cod_Cli")
            loComandoSeleccionar.AppendLine("INTO #tmpMovPendientesFact")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("WHERE Cod_Tip = 'FACT' AND Status <> 'Pagado' AND Mon_Sal > 0 ")
            loComandoSeleccionar.AppendLine("	AND Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("	AND Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("GROUP BY Cod_Cli")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT COUNT(Documento)								AS Pendientes,")
            loComandoSeleccionar.AppendLine("		Cod_Cli											AS Cod_Cli")
            loComandoSeleccionar.AppendLine("INTO #tmpMovPendientesAdel")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("WHERE Cod_Tip = 'ADEL' AND Status <> 'Pagado' AND Mon_Sal > 0 ")
            loComandoSeleccionar.AppendLine("	AND Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("	AND Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("GROUP BY Cod_Cli")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE	#tmpMovimientos")
            loComandoSeleccionar.AppendLine("SET		Orden = M.Orden,")
            loComandoSeleccionar.AppendLine("		Mon_Sal = M.Mon_Deb-M.Mon_Hab,")
            loComandoSeleccionar.AppendLine("		Sal_Ini = M.Sal_Ini")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT	ROW_NUMBER() ")
            loComandoSeleccionar.AppendLine("						OVER (	PARTITION BY #tmpMovimientos.Cod_Cli ")
            loComandoSeleccionar.AppendLine("								ORDER BY #tmpMovimientos.Fec_Ini, (CASE WHEN #tmpMovimientos.Cod_Tip='' THEN 'zzzzzzzzz' ELSE #tmpMovimientos.Cod_Tip END ) ASC) AS Orden,")
            loComandoSeleccionar.AppendLine("					#tmpMovimientos.Tabla, #tmpMovimientos.Cod_Tip, #tmpMovimientos.Documento,")
            loComandoSeleccionar.AppendLine("                   ISNULL(SI.Sal_Ini, @lnCero) AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("					#tmpMovimientos.Mon_Deb AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("					#tmpMovimientos.Mon_Hab AS Mon_Hab")
            loComandoSeleccionar.AppendLine("			FROM	#tmpMovimientos			")
            loComandoSeleccionar.AppendLine("			LEFT JOIN (SELECT Cod_Cli, SUM(Sal_Ini) AS Sal_Ini FROM #tmpSaldoInicial GROUP BY Cod_Cli) AS SI")
            loComandoSeleccionar.AppendLine("				ON SI.Cod_Cli = #tmpMovimientos.Cod_Cli")
            loComandoSeleccionar.AppendLine("		) AS M		")
            loComandoSeleccionar.AppendLine("WHERE	M.Tabla = #tmpMovimientos.Tabla ")
            loComandoSeleccionar.AppendLine("	AND	M.Cod_Tip = #tmpMovimientos.Cod_Tip")
            loComandoSeleccionar.AppendLine("	AND	M.Documento = #tmpMovimientos.Documento")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	A.Orden, A.Tabla, A.Cod_Cli, A.Nom_Cli, A.Cod_Tip, A.Documento, A.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("       CASE WHEN LEN(A.Referencia) > 60")
            loComandoSeleccionar.AppendLine("            THEN CONCAT(SUBSTRING(A.Referencia,1,60),'...')")
            loComandoSeleccionar.AppendLine("            ELSE A.Referencia")
            loComandoSeleccionar.AppendLine("       END         AS Referencia,")
            loComandoSeleccionar.AppendLine("		A.Sal_Ini, A.Mon_Deb, A.Mon_Hab, SUM(B.Mon_Sal) +  A.Sal_Ini AS Sal_Doc,")
            loComandoSeleccionar.AppendLine("       @sp_FecIni AS Desde, @sp_FecFin	AS Hasta,")
            loComandoSeleccionar.AppendLine("		COALESCE((#tmpMovPendientesFact.Pendientes),0) AS PendientesFact,")
            loComandoSeleccionar.AppendLine("		COALESCE((#tmpMovPendientesAdel.Pendientes),0) AS PendientesAdel")
            loComandoSeleccionar.AppendLine("FROM	#tmpMovimientos AS A")
            loComandoSeleccionar.AppendLine("	JOIN #tmpMovimientos AS B ON B.Cod_Cli = A.Cod_Cli")
            loComandoSeleccionar.AppendLine("		AND B.Orden <= A.Orden")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpMovPendientesFact ON #tmpMovPendientesFact.Cod_Cli = A.Cod_Cli")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpMovPendientesAdel ON #tmpMovPendientesAdel.Cod_Cli = A.Cod_Cli")
            loComandoSeleccionar.AppendLine("GROUP BY A.Orden, A.Tabla, A.Cod_Cli, A.Nom_Cli, A.Cod_Tip, A.Documento, A.Fec_Ini, A.Referencia, A.Sal_Ini,")
            loComandoSeleccionar.AppendLine("		A.Mon_Deb, A.Mon_Hab, #tmpMovPendientesFact.Pendientes, #tmpMovPendientesAdel.Pendientes")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  0 AS Orden, 'Saldo Inicial' AS Tabla, #tmpSaldoInicial.Cod_Cli, Clientes.Nom_Cli AS Nom_Cli, 'NR' AS Cod_Tip, ")
            loComandoSeleccionar.AppendLine("		@lcVacio AS Documento, @lcVacio AS Fec_Ini, 'SIN MOVIMIENTOS' AS Referencia, SUM(Sal_Ini) AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		@lnCero AS Mon_Deb, @lnCero AS Mon_Hab, @lnCero AS Sal_Doc,")
            loComandoSeleccionar.AppendLine("		@sp_FecIni AS Desde, @sp_FecFin	AS Hasta, ")
            loComandoSeleccionar.AppendLine("	    COALESCE(#tmpMovPendientesFact.Pendientes, 0) AS PendientesFact,")
            loComandoSeleccionar.AppendLine("	    COALESCE(#tmpMovPendientesAdel.Pendientes, 0) AS PendientesAdel")
            loComandoSeleccionar.AppendLine("FROM #tmpSaldoInicial")
            loComandoSeleccionar.AppendLine("  JOIN Clientes ON #tmpSaldoInicial.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpMovPendientesFact ON #tmpMovPendientesFact.Cod_Cli = #tmpSaldoInicial.Cod_Cli")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpMovPendientesAdel ON #tmpMovPendientesAdel.Cod_Cli = #tmpSaldoInicial.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE #tmpSaldoInicial.Cod_Cli NOT IN (SELECT Cod_Cli FROM #tmpMovimientos)")
            loComandoSeleccionar.AppendLine("GROUP BY #tmpSaldoInicial.Cod_Cli, Clientes.Nom_Cli, #tmpMovPendientesFact.Pendientes, #tmpMovPendientesAdel.Pendientes")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Cli ASC, Fec_Ini ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpSaldoInicial")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovPendientesFact")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovPendientesAdel")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rAnalitico_Clientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvTRF_rAnalitico_Clientes.ReportSource = loObjetoReporte

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
' DLC: 02/09/2010: Programacion inicial (Replica del reporte rLEstadoCuenta_HistoricoVentas)'
'                   - Cambio de la consulta a procedimiento almacenado.						'
'-------------------------------------------------------------------------------------------'
' DLC: 15/09/2010: Ajuste en la forma de obtener los detalles de Pagos, asi como también,	'
'                ajustar en el RPT, la forma de mostrar los detalles de Pagos.				'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Reprogramación del Reporte y su respectivo Store Procedure					'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Ajuste de la vista de Diseño.												'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Se elimino el filtro Detalle												'
'-------------------------------------------------------------------------------------------'
' RJG: 05/12/11: Eliminado el SP: ahora la consulta se hace desde un Query en línea para	'
'				 corregir cálculo de saldo y optimizar.										'
'-------------------------------------------------------------------------------------------'
