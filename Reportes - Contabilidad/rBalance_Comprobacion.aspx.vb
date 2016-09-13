'-------------------------------------------------------------------------------------------'
' Inicio del codigo																			'
'-------------------------------------------------------------------------------------------'
' Importando librerias																		'
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rBalance_Comprobacion"													'
'-------------------------------------------------------------------------------------------'
Partial Class rBalance_Comprobacion
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
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))

            Dim llSoloMovimientos As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(7)).Trim().ToUpper().Equals("SI")

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Desde DATETIME	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Hasta DATETIME	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Desde VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Hasta VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Desde VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Hasta VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Desde VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Hasta VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Desde VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Hasta VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro6Hasta VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro6Desde VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro7Desde VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro8Desde VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("DECLARE @lcParametro9Desde VARCHAR(100)	")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET		@lcParametro1Desde = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("SET		@lcParametro1Hasta = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("SET		@lcParametro2Desde = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("SET		@lcParametro2Hasta = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("SET		@lcParametro3Desde = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("SET		@lcParametro3Hasta = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("SET		@lcParametro4Desde = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("SET		@lcParametro4Hasta = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("SET		@lcParametro5Desde = " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("SET		@lcParametro5Hasta = " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("SET		@lcParametro6Desde = " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("SET		@lcParametro6Hasta = " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("SET		@lcParametro7Desde = " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("SET		@lcParametro8Desde = " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("SET		@lcParametro9Desde = " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10)")
            loComandoSeleccionar.AppendLine("DECLARE @llFalso BIT")
            loComandoSeleccionar.AppendLine("DECLARE @llVerdadero BIT")
            loComandoSeleccionar.AppendLine("SET	@lnCero = 0")
            loComandoSeleccionar.AppendLine("SET	@llFalso = 0")
            loComandoSeleccionar.AppendLine("SET	@llVerdadero = 1")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("-- Prepara un listado de las cuentas contables a incluir  *")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("SELECT		LEN(RTRIM(Cod_Cue))				AS Nivel, ")
            loComandoSeleccionar.AppendLine("			CAST(Cod_Cue AS VARCHAR(100))	AS Cod_cue, ")
            loComandoSeleccionar.AppendLine("			CAST(Nom_Cue AS VARCHAR(100))	AS Nom_Cue, ")
            loComandoSeleccionar.AppendLine("			Movimiento						AS Movimiento,")
            loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_1,")
            loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_2,")
            loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_3,")
            loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_4,")
            loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_5,")
            loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_6,")
            loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_7,")
            loComandoSeleccionar.AppendLine("			Mon_Ini							AS Saldo_Inicial")
            loComandoSeleccionar.AppendLine("INTO		#tmpCuentas")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Contables")
            loComandoSeleccionar.AppendLine("WHERE		Cod_Cue BETWEEN @lcParametro2Desde AND @lcParametro2Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("-- Agrega los saldos iniciales.							  *")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("UPDATE		#tmpCuentas")
            loComandoSeleccionar.AppendLine("SET		#tmpCuentas.Saldo_Inicial = #tmpCuentas.Saldo_Inicial ")
            loComandoSeleccionar.AppendLine("										+ Saldos.TotalDebe ")
            loComandoSeleccionar.AppendLine("										- Saldos.Total_Haber")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT	RC.Cod_Cue			AS Cod_cue, ")
            loComandoSeleccionar.AppendLine("					SUM(RC.Mon_Deb)		AS TotalDebe,")
            loComandoSeleccionar.AppendLine("					SUM(RC.Mon_Hab)		AS Total_Haber")
            loComandoSeleccionar.AppendLine("			FROM	Renglones_Comprobantes AS RC")
            loComandoSeleccionar.AppendLine("				JOIN Comprobantes ON Comprobantes.Documento = RC.Documento")
            loComandoSeleccionar.AppendLine("					AND Comprobantes.Adicional = RC.Adicional ")
            loComandoSeleccionar.AppendLine("					AND Comprobantes.Tipo = @lcParametro7Desde ")
            loComandoSeleccionar.AppendLine("					AND	Comprobantes.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			WHERE	RC.Fec_Ini < @lcParametro1Desde")
            loComandoSeleccionar.AppendLine("			AND RC.Cod_Cen	BETWEEN @lcParametro3Desde AND @lcParametro3Hasta")
            loComandoSeleccionar.AppendLine("			AND RC.Cod_Gas	BETWEEN @lcParametro4Desde AND @lcParametro4Hasta")
            loComandoSeleccionar.AppendLine("			AND RC.Cod_Aux	BETWEEN @lcParametro5Desde AND @lcParametro5Hasta")
            loComandoSeleccionar.AppendLine("			AND RC.Cod_Mon	BETWEEN @lcParametro6Desde AND @lcParametro6Hasta")
            loComandoSeleccionar.AppendLine("			GROUP BY RC.Cod_Cue")
            loComandoSeleccionar.AppendLine("		) AS Saldos")
            loComandoSeleccionar.AppendLine("WHERE	#tmpCuentas.Cod_Cue = Saldos.Cod_Cue")
            loComandoSeleccionar.AppendLine("	AND #tmpCuentas.Movimiento = 1")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("-- Calcula el 'rango' de cada nivel.                      *")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("UPDATE		#tmpCuentas")
            loComandoSeleccionar.AppendLine("SET		Nivel_1 = (CASE WHEN N.Margen > 1 THEN 1 ELSE 0 END),")
            loComandoSeleccionar.AppendLine("			Nivel_2 = (CASE WHEN N.Margen > 2 THEN 1 ELSE 0 END),")
            loComandoSeleccionar.AppendLine("			Nivel_3 = (CASE WHEN N.Margen > 3 THEN 1 ELSE 0 END),")
            loComandoSeleccionar.AppendLine("			Nivel_4 = (CASE WHEN N.Margen > 4 THEN 1 ELSE 0 END),")
            loComandoSeleccionar.AppendLine("			Nivel_5 = (CASE WHEN N.Margen > 5 THEN 1 ELSE 0 END),")
            loComandoSeleccionar.AppendLine("			Nivel_6 = (CASE WHEN N.Margen > 6 THEN 1 ELSE 0 END),")
            loComandoSeleccionar.AppendLine("			Nivel_7 = (CASE WHEN N.Margen > 7 THEN 1 ELSE 0 END),")
            loComandoSeleccionar.AppendLine("			Nivel = N.Margen			")
            loComandoSeleccionar.AppendLine("FROM		(	SELECT	ROW_NUMBER() OVER(ORDER BY #tmpCuentas.Nivel ASC) AS Margen,")
            loComandoSeleccionar.AppendLine("						#tmpCuentas.Nivel AS Original")
            loComandoSeleccionar.AppendLine("				FROM	#tmpCuentas")
            loComandoSeleccionar.AppendLine("				GROUP BY #tmpCuentas.Nivel")
            loComandoSeleccionar.AppendLine("			) AS N")
            loComandoSeleccionar.AppendLine("WHERE		#tmpCuentas.Nivel = N.Original")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("-- Union principal.                                       *")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("SELECT		#tmpCuentas.Nivel								AS Nivel,")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Cod_Cue								AS Cod_Cue, ")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Nom_Cue								AS Nom_Cue, ")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Movimiento							AS Movimiento,")
            loComandoSeleccionar.AppendLine("			SUM(ISNULL(tmpRenglones.Mon_Deb, @lnCero))		AS Mon_Deb, ")
            loComandoSeleccionar.AppendLine("			SUM(ISNULL(tmpRenglones.Mon_Hab, @lnCero))		AS Mon_Hab, ")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_1								AS Nivel_1,")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_2								AS Nivel_2,")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_3								AS Nivel_3,")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_4								AS Nivel_4,")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_5								AS Nivel_5,")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_6								AS Nivel_6,")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_7								AS Nivel_7,")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Saldo_Inicial						AS Saldo_Inicial,")
            loComandoSeleccionar.AppendLine("			(	SUM(ISNULL(tmpRenglones.Mon_Deb, @lnCero)) -				")
            loComandoSeleccionar.AppendLine("				SUM(ISNULL(tmpRenglones.Mon_Hab, @lnCero)) +				")
            loComandoSeleccionar.AppendLine("				#tmpCuentas.Saldo_Inicial)					AS Saldo_Actual")
            loComandoSeleccionar.AppendLine("INTO		#tmpTodos")
            loComandoSeleccionar.AppendLine("FROM		#tmpCuentas")
            loComandoSeleccionar.AppendLine("	LEFT JOIN (")
            loComandoSeleccionar.AppendLine("				SELECT	Renglones_Comprobantes.Mon_Deb,")
            loComandoSeleccionar.AppendLine("						Renglones_Comprobantes.Mon_Hab,")
            loComandoSeleccionar.AppendLine("						Renglones_Comprobantes.Cod_cue")
            loComandoSeleccionar.AppendLine("				FROM	Renglones_Comprobantes")
            loComandoSeleccionar.AppendLine("					JOIN	Comprobantes")
            loComandoSeleccionar.AppendLine("						ON	Comprobantes.Documento = Renglones_Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("						AND Comprobantes.Adicional = Renglones_Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine("						AND Comprobantes.Tipo = @lcParametro7Desde")
            loComandoSeleccionar.AppendLine("					AND Comprobantes.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("				WHERE		Renglones_Comprobantes.Fec_Ini	BETWEEN @lcParametro1Desde AND @lcParametro1Hasta")
            loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Cen	BETWEEN @lcParametro3Desde AND @lcParametro3Hasta")
            loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Gas	BETWEEN @lcParametro4Desde AND @lcParametro4Hasta")
            loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Aux	BETWEEN @lcParametro5Desde AND @lcParametro5Hasta")
            loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Mon	BETWEEN @lcParametro6Desde AND @lcParametro6Hasta")
            loComandoSeleccionar.AppendLine("			) AS tmpRenglones ")
            loComandoSeleccionar.AppendLine("		ON tmpRenglones.Cod_Cue = #tmpCuentas.Cod_Cue")
            loComandoSeleccionar.AppendLine("GROUP BY	#tmpCuentas.Nivel, #tmpCuentas.Cod_cue, #tmpCuentas.Nom_Cue, ")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Movimiento, #tmpCuentas.Saldo_Inicial,")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_1, #tmpCuentas.Nivel_2, #tmpCuentas.Nivel_3, ")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_4, #tmpCuentas.Nivel_5, #tmpCuentas.Nivel_6, ")
            loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_7")
            loComandoSeleccionar.AppendLine("ORDER BY #tmpCuentas.Cod_cue")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCuentas")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("-- Actualiza los totales hasta el nivel indicado.         *")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE	#tmpTodos")
            loComandoSeleccionar.AppendLine("SET	Mon_Deb = T.Total_Debe,")
            loComandoSeleccionar.AppendLine("		Mon_Hab = T.Total_Haber,")
            loComandoSeleccionar.AppendLine("		Saldo_Inicial = T.Total_Inicial,")
            loComandoSeleccionar.AppendLine("		Saldo_Actual = T.Total_Actual")
            loComandoSeleccionar.AppendLine("FROM (	SELECT	Total.Cod_cue				AS Cuenta_Padre,")
            loComandoSeleccionar.AppendLine("				SUM(Interna.Mon_Deb)		AS Total_Debe,")
            loComandoSeleccionar.AppendLine("				SUM(Interna.Mon_Hab)		AS Total_Haber,")
            loComandoSeleccionar.AppendLine("				SUM(Interna.Saldo_Inicial)	AS Total_Inicial,")
            loComandoSeleccionar.AppendLine("				SUM(Interna.Saldo_Actual)	AS Total_Actual")
            loComandoSeleccionar.AppendLine("		FROM	#tmpTodos AS Total")
            loComandoSeleccionar.AppendLine("			JOIN #tmpTodos AS Interna")
            loComandoSeleccionar.AppendLine("				ON SUBSTRING(Interna.Cod_Cue, 1, LEN(Total.Cod_Cue)) = Total.Cod_cue")
            loComandoSeleccionar.AppendLine("				AND Interna.Nivel >= @lcParametro9Desde")
            loComandoSeleccionar.AppendLine("		GROUP BY Total.Cod_cue")
            loComandoSeleccionar.AppendLine("	) AS T")
            loComandoSeleccionar.AppendLine("WHERE #tmpTodos.Nivel = @lcParametro9Desde")
            loComandoSeleccionar.AppendLine("	AND T.Cuenta_Padre = #tmpTodos.Cod_Cue")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("-- Actualiza el máximo nivel (lo requiere el RPT)         *")
            loComandoSeleccionar.AppendLine("--*********************************************************")
            loComandoSeleccionar.AppendLine("DELETE	FROM #tmpTodos")
            loComandoSeleccionar.AppendLine("WHERE	#tmpTodos.Nivel > @lcParametro9Desde")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET @lcParametro9Desde = (SELECT MAX(Nivel) FROM #tmpTodos)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            If llSoloMovimientos Then
                loComandoSeleccionar.AppendLine("--*********************************************************")
                loComandoSeleccionar.AppendLine("-- Elimina los niveles máximos y sin movimientos          *")
                loComandoSeleccionar.AppendLine("--*********************************************************")
                loComandoSeleccionar.AppendLine("DELETE		FROM #tmpTodos")
                loComandoSeleccionar.AppendLine("WHERE		#tmpTodos.Nivel = @lcParametro9Desde")
                loComandoSeleccionar.AppendLine("	AND		(ABS(Mon_Deb)+ABS(Mon_Hab)+ABS(Saldo_Inicial)+ABS(Saldo_Actual)) = @lnCero")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("--*********************************************************")
                loComandoSeleccionar.AppendLine("-- Elimina los niveles superiores y sin movimientos       *")
                loComandoSeleccionar.AppendLine("--*********************************************************")
                loComandoSeleccionar.AppendLine("SELECT	@lnCero As SubTotal, Cod_Cue, Nivel")
                loComandoSeleccionar.AppendLine("INTO	#tmpTotales")
                loComandoSeleccionar.AppendLine("FROM	#tmpTodos")
                loComandoSeleccionar.AppendLine("WHERE	#tmpTodos.Nivel < @lcParametro9Desde")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("DECLARE curNiveles CURSOR FOR")
                loComandoSeleccionar.AppendLine("	SELECT DISTINCT Nivel FROM #tmpTotales ORDER BY Nivel DESC")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("OPEN curNiveles ")
                loComandoSeleccionar.AppendLine("DECLARE @lnNivelActual INT")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("FETCH NEXT FROM curNiveles INTO @lnNivelActual")
                loComandoSeleccionar.AppendLine("WHILE @@FETCH_STATUS = 0")
                loComandoSeleccionar.AppendLine("BEGIN")
                loComandoSeleccionar.AppendLine("		UPDATE		#tmpTotales")
                loComandoSeleccionar.AppendLine("		SET			#tmpTotales.SubTotal = SubTotales.SubTotal")
                loComandoSeleccionar.AppendLine("		FROM		(	SELECT		(SUM(Mon_Deb)+SUM(Mon_Hab)+ABS(SUM(Saldo_Inicial))+ABS(SUM(Saldo_Actual))) AS SubTotal,")
                loComandoSeleccionar.AppendLine("									B.Cod_Cue")
                loComandoSeleccionar.AppendLine("						FROM		#tmpTodos AS A")
                loComandoSeleccionar.AppendLine("							JOIN	#tmpTotales AS B ON A.Cod_Cue LIKE (RTRIM(B.Cod_Cue) + '%')")
                loComandoSeleccionar.AppendLine("						WHERE		A.Nivel > B.Nivel")
                loComandoSeleccionar.AppendLine("							AND		B.Nivel = @lnNivelActual")
                loComandoSeleccionar.AppendLine("						GROUP BY	B.Cod_Cue")
                loComandoSeleccionar.AppendLine("					) AS SubTotales")
                loComandoSeleccionar.AppendLine("		WHERE		SubTotales.Cod_Cue = #tmpTotales.Cod_Cue")
                loComandoSeleccionar.AppendLine("			AND		#tmpTotales.Nivel = @lnNivelActual")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("		FETCH NEXT FROM curNiveles INTO @lnNivelActual")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("END")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("CLOSE curNiveles ")
                loComandoSeleccionar.AppendLine("DEALLOCATE curNiveles ")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("DELETE #tmpTodos")
                loComandoSeleccionar.AppendLine("FROM	#tmpTotales As Totales")
                loComandoSeleccionar.AppendLine("WHERE	#tmpTodos.Cod_cue = Totales.Cod_Cue")
                loComandoSeleccionar.AppendLine("		AND Totales.SubTotal = @lnCero")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("DROP TABLE #tmpTotales")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("")
            End If
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		CAST(@lcParametro9Desde AS INT) AS Nivel_Maximo, ")
            loComandoSeleccionar.AppendLine("			#tmpTodos.* ")
            loComandoSeleccionar.AppendLine("FROM		#tmpTodos")
            loComandoSeleccionar.AppendLine("ORDER BY	Cod_cue")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpTodos")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes", 1800)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rBalance_Comprobacion", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrBalance_Comprobacion.ReportSource = loObjetoReporte

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
' RJG: 26/10/11: Codigo inicial, a partir de Libro Diario.									'
'-------------------------------------------------------------------------------------------' 
' RJG: 27/10/11: Corrección en cálculo de saldos inicial y final. Corrección en filtro.		'
'-------------------------------------------------------------------------------------------' 
' RJG: 16/11/11: Agregado filtro de la estructura superior cuando en el detalle no hay		'
'				 movimientos (y el usuario indicó el filtro "Solo Movimientos = SI".		'
'-------------------------------------------------------------------------------------------' 
' RJG: 06/12/11: Se agregó la igualdad de campo Adicional en las uniones entre Comprobantes	'
'				 y sus renglones.															'
'-------------------------------------------------------------------------------------------' 
' RJG: 09/02/12: Cambiado el filtro de comprobantes "='Pendiente'" por "<>'Anulado'".		'
'-------------------------------------------------------------------------------------------' 
' JFP: 04/07/13: Ampliacion del tiempo para el timeout                                      '
'-------------------------------------------------------------------------------------------' 
