﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rAnalitico_Clientes"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rAnalitico_Clientes
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

            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecIni			AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecFin			AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_CodCli_Desde	AS VARCHAR(10)")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_CodCli_Hasta	AS VARCHAR(10)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET	@sp_FecIni          = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("SET	@sp_FecFin          = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("SET	@sp_CodCli_Desde    = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("SET	@sp_CodCli_Hasta    = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero	AS DECIMAL(28, 10)	;")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio	AS VARCHAR(10)	;")
            loComandoSeleccionar.AppendLine("SET @lnCero		= 0")
            loComandoSeleccionar.AppendLine("SET @lcVacio		= ''")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
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
            loComandoSeleccionar.AppendLine("WHERE	Cuentas_Cobrar.Fec_Ini < @sp_FecIni")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("GROUP BY Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("		Cobros.Cod_Cli,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Cobros.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN Renglones_Cobros.Mon_Abo")
            loComandoSeleccionar.AppendLine("				ELSE -Renglones_Cobros.Mon_Abo	")
            loComandoSeleccionar.AppendLine("			END) -(Cobros.Mon_Ret - Cobros.Mon_Des) AS Sal_Ini")
            loComandoSeleccionar.AppendLine("FROM	Cobros")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Cobros ON Cobros.Documento = Renglones_Cobros.Documento		")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON Clientes.Cod_Cli = Cobros.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE	Cobros.Status IN ('Confirmado')")
            loComandoSeleccionar.AppendLine("		AND Cobros.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("		AND	Cobros.Fec_Ini < @sp_FecIni")
            loComandoSeleccionar.AppendLine("		AND Cobros.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("GROUP BY Cobros.Cod_Cli, Cobros.Mon_Ret, Cobros.Mon_Des")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--MOVIMIENTOS")
            loComandoSeleccionar.AppendLine("SELECT 'Cobros'										AS Tabla,")
            loComandoSeleccionar.AppendLine("		0											AS Orden,")
            loComandoSeleccionar.AppendLine("		Clientes.Cod_Cli							AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		Clientes.nom_cli							AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		'Cobro'	+'-'+ Renglones_Cobros.Cod_Tip		AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cobros.Documento							AS Documento,")
            loComandoSeleccionar.AppendLine("		Renglones_Cobros.Doc_Ori					AS Origen,")
            loComandoSeleccionar.AppendLine("		Cobros.Fec_Ini								AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		CONCAT(RTRIM(Detalles_Cobros.Tip_Ope),' ', ")
            loComandoSeleccionar.AppendLine("			RTRIM(Detalles_Cobros.Num_Doc), '/ ',")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Cobros.Comentario, 6, 30))    AS Referencia,")
            loComandoSeleccionar.AppendLine("		@lnCero										AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Cobros.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN	Renglones_Cobros.Mon_Abo	")
            loComandoSeleccionar.AppendLine("				ELSE	@lnCero")
            loComandoSeleccionar.AppendLine("		END)										AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Cobros.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN	@lnCero	")
            loComandoSeleccionar.AppendLine("				ELSE	Renglones_Cobros.Mon_Abo	")
            loComandoSeleccionar.AppendLine("		END) -(Cobros.Mon_Ret + Cobros.Mon_Des) 	AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("		@lnCero										AS Mon_Sal")
            loComandoSeleccionar.AppendLine("INTO #tmpDetallesCobros")
            loComandoSeleccionar.AppendLine("FROM Cobros")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Cobros ON Cobros.Documento = Renglones_Cobros.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Detalles_Cobros ON Cobros.Documento = Detalles_Cobros.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Clientes ON Cobros.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE	Cobros.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("		AND	Cobros.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("		AND Cobros.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta		")
            loComandoSeleccionar.AppendLine("GROUP BY Clientes.Cod_Cli, Clientes.Nom_Cli, Cobros.Documento, Renglones_Cobros.Doc_Ori, Cobros.Fec_Ini,Renglones_Cobros.Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cobros.Comentario,Cobros.Mon_Ret, Cobros.Mon_Des,Detalles_Cobros.Tip_Ope,Detalles_Cobros.Num_Doc")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT 'Cobros'										AS Tabla,")
            loComandoSeleccionar.AppendLine("		0											AS Orden,")
            loComandoSeleccionar.AppendLine("		Clientes.Cod_Cli							AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		Clientes.nom_cli							AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		'Cobro' +'-'+ Renglones_Cobros.Cod_Tip		AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cobros.Documento							AS Documento,")
            loComandoSeleccionar.AppendLine("		Renglones_Cobros.Doc_Ori					AS Origen,")
            loComandoSeleccionar.AppendLine("		Cobros.Fec_Ini								AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		CASE WHEN Renglones_Cobros.Cod_Tip = 'ADEL' THEN")
            loComandoSeleccionar.AppendLine("            (SELECT CONCAT(RTRIM(Tip_Ope),' ', RTRIM(Num_Doc))")
            loComandoSeleccionar.AppendLine("            FROM Detalles_Cobros")
            loComandoSeleccionar.AppendLine("				JOIN Renglones_Cobros AS RC ON Detalles_Cobros.Documento = RC.Documento")
            loComandoSeleccionar.AppendLine("            WHERE RC.Doc_Ori = Renglones_Cobros.Doc_Ori AND RC.Cod_Tip = 'ADEL') ")
            loComandoSeleccionar.AppendLine("            ELSE Renglones_Cobros.Comentario")
            loComandoSeleccionar.AppendLine("        END												AS Referencia,")
            loComandoSeleccionar.AppendLine("		@lnCero										AS Sal_Ini,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Cobros.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN	Renglones_Cobros.Mon_Abo	")
            loComandoSeleccionar.AppendLine("				ELSE	@lnCero")
            loComandoSeleccionar.AppendLine("		END)										AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("		SUM(CASE WHEN Renglones_Cobros.Tip_Doc = 'Debito' ")
            loComandoSeleccionar.AppendLine("				THEN	@lnCero	")
            loComandoSeleccionar.AppendLine("				ELSE	Renglones_Cobros.Mon_Abo	")
            loComandoSeleccionar.AppendLine("		END) -(Cobros.Mon_Ret + Cobros.Mon_Des) 	AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("		@lnCero										AS Mon_Sal")
            loComandoSeleccionar.AppendLine("INTO #tmpNoDetallesCobros")
            loComandoSeleccionar.AppendLine("FROM Cobros")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Cobros ON Cobros.Documento = Renglones_Cobros.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Clientes ON Cobros.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE	Cobros.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("		AND	Cobros.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("		AND Cobros.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("		AND Renglones_Cobros.Doc_Ori NOT IN (SELECT Origen FROM #tmpDetallesCobros)")
            loComandoSeleccionar.AppendLine("GROUP BY Clientes.Cod_Cli, Clientes.Nom_Cli, Cobros.Documento, Renglones_Cobros.Doc_Ori, Renglones_Cobros.Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cobros.Fec_Ini, Cobros.Comentario,Cobros.Mon_Ret, Cobros.Mon_Des, Renglones_Cobros.Comentario")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT * INTO #tmpCobros FROM #tmpNoDetallesCobros")
            loComandoSeleccionar.AppendLine("UNION")
            loComandoSeleccionar.AppendLine("SELECT * FROM #tmpDetallesCobros")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Cuentas_Cobrar'								AS Tabla,")
            loComandoSeleccionar.AppendLine("		0												AS Orden,")
            loComandoSeleccionar.AppendLine("		Clientes.Cod_Cli								AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		Clientes.Nom_Cli								AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Cod_Tip							AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Documento						AS Documento,")
            loComandoSeleccionar.AppendLine("		@lcVacio										AS Origen,")
            loComandoSeleccionar.AppendLine("		Cuentas_Cobrar.Fec_Ini							AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		CASE WHEN Cuentas_Cobrar.Cod_Tip = 'ADEL' THEN")
            loComandoSeleccionar.AppendLine("            (SELECT CONCAT(RTRIM(Tip_Ope),' ', RTRIM(Num_Doc))")
            loComandoSeleccionar.AppendLine("            FROM Detalles_Cobros")
            loComandoSeleccionar.AppendLine("            WHERE Documento = Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("            ) ")
            loComandoSeleccionar.AppendLine("            ELSE Cuentas_Cobrar.Comentario")
            loComandoSeleccionar.AppendLine("        END												AS Referencia,")
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
            loComandoSeleccionar.AppendLine("INTO #tmpCuentasCobrar")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes  ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE	Cuentas_Cobrar.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Documento NOT IN (SELECT Origen FROM #tmpCobros)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT * INTO #tmpMovimientos")
            loComandoSeleccionar.AppendLine("FROM #tmpCobros")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("SELECT * FROM #tmpCuentasCobrar")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--PENDIENTES POR CONCILIAR")
            loComandoSeleccionar.AppendLine("SELECT COUNT(Documento)									AS Pendientes,")
            loComandoSeleccionar.AppendLine("		Cod_Cli											AS Cod_Cli")
            loComandoSeleccionar.AppendLine("INTO #tmpMovPendientes")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("WHERE Cod_Tip = 'FACT' AND Status <> 'Pagado' AND Mon_Sal > 0 ")
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
            loComandoSeleccionar.AppendLine("SELECT	A.Orden, A.Tabla, A.Cod_Cli, A.Nom_Cli, A.Cod_Tip, A.Documento, A.Fec_Ini, A.Referencia, A.Sal_Ini,A.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("		A.Mon_Hab, SUM(B.Mon_Sal) +  A.Sal_Ini AS Sal_Doc, @sp_FecIni AS Desde, @sp_FecFin	AS Hasta,")
            loComandoSeleccionar.AppendLine("        ISNULL((SELECT Pendientes FROM #tmpMovPendientes WHERE Cod_Cli = A.Cod_Cli),0) AS Pendientes")
            loComandoSeleccionar.AppendLine("FROM	#tmpMovimientos AS A")
            loComandoSeleccionar.AppendLine("	JOIN #tmpMovimientos AS B ON B.Cod_Cli = A.Cod_Cli")
            loComandoSeleccionar.AppendLine("		AND B.Orden <= A.Orden")
            loComandoSeleccionar.AppendLine("GROUP BY A.Orden, A.Tabla, A.Cod_Cli, A.Nom_Cli, A.Cod_Tip, A.Documento, A.Fec_Ini, A.Referencia, A.Sal_Ini,")
            loComandoSeleccionar.AppendLine("		A.Mon_Deb, A.Mon_Hab")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Cli ASC, A.Fec_Ini ASC, (CASE WHEN A.Cod_Tip='' THEN 'zzzzzzzzz' ELSE A.Cod_Tip END ) ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpSaldoInicial")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovPendientes")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpDetallesCobros")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpNoDetallesCobros")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCobros")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCuentasCobrar")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then

                Dim loComandoSeleccionar1 As New StringBuilder()

                loComandoSeleccionar1.AppendLine("DECLARE	@sp_FecIni			AS DATETIME")
                loComandoSeleccionar1.AppendLine("DECLARE	@sp_FecFin			AS DATETIME")
                loComandoSeleccionar1.AppendLine("DECLARE	@sp_CodCli_Desde	AS VARCHAR(10)")
                loComandoSeleccionar1.AppendLine("DECLARE	@sp_CodCli_Hasta	AS VARCHAR(10)")
                loComandoSeleccionar1.AppendLine("")
                loComandoSeleccionar1.AppendLine("SET	@sp_FecIni          = " & lcParametro0Desde)
                loComandoSeleccionar1.AppendLine("SET	@sp_FecFin          = " & lcParametro0Hasta)
                loComandoSeleccionar1.AppendLine("SET	@sp_CodCli_Desde    = " & lcParametro1Desde)
                loComandoSeleccionar1.AppendLine("SET	@sp_CodCli_Hasta    = " & lcParametro1Hasta)
                loComandoSeleccionar1.AppendLine("")
                loComandoSeleccionar1.AppendLine("--SALDO INICIAL")
                loComandoSeleccionar1.AppendLine("SELECT	Cuentas_Cobrar.Cod_Cli							AS Cod_Cli,")
                loComandoSeleccionar1.AppendLine("		SUM(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Debito' ")
                loComandoSeleccionar1.AppendLine("				THEN Cuentas_Cobrar.Mon_Net")
                loComandoSeleccionar1.AppendLine("				ELSE -Cuentas_Cobrar.Mon_Net	")
                loComandoSeleccionar1.AppendLine("			END)										AS Sal_Ini")
                loComandoSeleccionar1.AppendLine("INTO #tmpSaldoInicial")
                loComandoSeleccionar1.AppendLine("FROM	Cuentas_Cobrar")
                loComandoSeleccionar1.AppendLine("	JOIN Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
                loComandoSeleccionar1.AppendLine("WHERE	Cuentas_Cobrar.Fec_Ini < @sp_FecIni")
                loComandoSeleccionar1.AppendLine("		AND Cuentas_Cobrar.Status <> 'Anulado'")
                loComandoSeleccionar1.AppendLine("			AND Cuentas_Cobrar.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
                loComandoSeleccionar1.AppendLine("GROUP BY Cuentas_Cobrar.Cod_Cli")
                loComandoSeleccionar1.AppendLine("")
                loComandoSeleccionar1.AppendLine("UNION ALL")
                loComandoSeleccionar1.AppendLine("")
                loComandoSeleccionar1.AppendLine("SELECT")
                loComandoSeleccionar1.AppendLine("		Cobros.Cod_Cli,")
                loComandoSeleccionar1.AppendLine("		SUM(CASE WHEN Renglones_Cobros.Tip_Doc = 'Debito' ")
                loComandoSeleccionar1.AppendLine("				THEN Renglones_Cobros.Mon_Abo")
                loComandoSeleccionar1.AppendLine("				ELSE -Renglones_Cobros.Mon_Abo	")
                loComandoSeleccionar1.AppendLine("			END) -(Cobros.Mon_Ret - Cobros.Mon_Des) AS Sal_Ini")
                loComandoSeleccionar1.AppendLine("FROM	Cobros")
                loComandoSeleccionar1.AppendLine("	JOIN	Renglones_Cobros ON Cobros.Documento = Renglones_Cobros.Documento		")
                loComandoSeleccionar1.AppendLine("	JOIN	Clientes ON Clientes.Cod_Cli = Cobros.Cod_Cli")
                loComandoSeleccionar1.AppendLine("WHERE	Cobros.Status = 'Confirmado'")
                loComandoSeleccionar1.AppendLine("		AND Cobros.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
                loComandoSeleccionar1.AppendLine("		AND	Cobros.Fec_Ini < @sp_FecIni")
                loComandoSeleccionar1.AppendLine("		AND Cobros.Cod_Cli BETWEEN @sp_CodCli_Desde AND @sp_CodCli_Hasta")
                loComandoSeleccionar1.AppendLine("GROUP BY Cobros.Cod_Cli, Cobros.Mon_Ret, Cobros.Mon_Des")
                loComandoSeleccionar1.AppendLine("")
                loComandoSeleccionar1.AppendLine("SELECT	'' AS Orden, '' AS Tabla, #tmpSaldoInicial.Cod_Cli, Clientes.Nom_Cli AS Nom_Cli, 'NR' AS Cod_Tip, ")
                loComandoSeleccionar1.AppendLine("		'' AS Documento, '' AS Fec_Ini, '' AS Referencia, SUM(Sal_Ini) AS Sal_Ini,")
                loComandoSeleccionar1.AppendLine("		0 AS Mon_Deb,0 AS Mon_Hab, 0 AS Sal_Doc, 0 AS Pendientes,")
                loComandoSeleccionar1.AppendLine("       @sp_FecIni AS Desde, @sp_FecFin	AS Hasta")
                loComandoSeleccionar1.AppendLine("FROM #tmpSaldoInicial")
                loComandoSeleccionar1.AppendLine("  JOIN Clientes ON #tmpSaldoInicial.Cod_Cli = Clientes.Cod_Cli")
                loComandoSeleccionar1.AppendLine("GROUP BY #tmpSaldoInicial.Cod_Cli,Clientes.Nom_Cli")
                loComandoSeleccionar1.AppendLine("DROP TABLE #tmpSaldoInicial")

                'Me.mEscribirConsulta(loComandoSeleccionar1.ToString())
                Dim loServicios1 As New cusDatos.goDatos
                Dim laDatosReporte1 As DataSet = loServicios1.mObtenerTodosSinEsquema(loComandoSeleccionar1.ToString(), "curReportes")

                Me.mCargarLogoEmpresa(laDatosReporte1.Tables(0), "LogoEmpresa")

                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rAnalitico_Clientes", laDatosReporte1)

                Me.mTraducirReporte(loObjetoReporte)
                Me.mFormatearCamposReporte(loObjetoReporte)
                Me.crvCGS_rAnalitico_Clientes.ReportSource = loObjetoReporte

                If (laDatosReporte1.Tables(0).Rows.Count <= 0) Then

                    Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                              "No se Encontraron Registros para los Parámetros Especificados. ", _
                                               vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                               "350px", _
                                               "200px")
                End If
            Else
                loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rAnalitico_Clientes", laDatosReporte)

                Me.mTraducirReporte(loObjetoReporte)
                Me.mFormatearCamposReporte(loObjetoReporte)
                Me.crvCGS_rAnalitico_Clientes.ReportSource = loObjetoReporte
            End If

            'loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rAnalitico_Clientes", laDatosReporte)

            'Me.mTraducirReporte(loObjetoReporte)
            'Me.mFormatearCamposReporte(loObjetoReporte)
            'Me.crvCGS_rAnalitico_Clientes.ReportSource = loObjetoReporte

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