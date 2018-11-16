'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rBalance_General"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rBalance_General

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcFechaDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcFechaHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcCuentaContableDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcCuentaContableHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcAuxiliarDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcAuxiliarHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim llSoloMovimientos As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(3)).Trim().ToUpper().Equals("SI")
            Dim lnNivelMax As Integer = CInt(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcMostrarAux As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lnLongMax As Integer = 0
            Select Case lnNivelMax
                Case 1
                    lnLongMax = 1
                Case 2
                    lnLongMax = 3
                Case 3
                    lnLongMax = 5
                Case 4
                    lnLongMax = 8
                Case 5
                    lnLongMax = 12
            End Select

            Dim loComandoSeleccionar As New StringBuilder()
            
            loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10) = 0")
            loComandoSeleccionar.AppendLine("DECLARE @lcFechaDesde DATETIME = " & lcFechaDesde)
            loComandoSeleccionar.AppendLine("DECLARE @lcFechaHasta DATETIME = " & lcFechaHasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCuentaDesde VARCHAR(30) = " & lcCuentaContableDesde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCuentaHasta VARCHAR(30) = " & lcCuentaContableHasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcAuxDesde VARCHAR(30) = " & lcAuxiliarDesde)
            loComandoSeleccionar.AppendLine("DECLARE @lcAuxHasta VARCHAR(30) = " & lcAuxiliarHasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcMostrarAux VARCHAR(12) = '" & lcMostrarAux & "'")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnNivelMax INT = " & lnNivelMax)
            loComandoSeleccionar.AppendLine("DECLARE @lnLongMax INT = " & lnLongMax)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	SUBSTRING(CC.Cod_Cue, 1, @lnLongMax)					AS Cod_Cue,")
            loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>1 AND @lnNivelMax > 1) THEN 1 ELSE 0 END)	AS Cod_Niv_1,")
            loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>3 AND @lnNivelMax > 2) THEN 3 ELSE 0 END)	AS Cod_Niv_2,")
            loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>5 AND @lnNivelMax > 3) THEN 5 ELSE 0 END)	AS Cod_Niv_3,")
            loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>8 AND @lnNivelMax > 4) THEN 8 ELSE 0 END)	AS Cod_Niv_4,")
            loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>12 AND @lnNivelMax > 5) THEN 12 ELSE 0 END)	AS Cod_Niv_5,")
            loComandoSeleccionar.AppendLine("		COALESCE(Renglones_Comprobantes.Cod_Aux,'') AS Auxiliar,")
            loComandoSeleccionar.AppendLine("       COALESCE(Auxiliares.Nom_Aux, '')			AS Nom_Auxiliar,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < @lcFechaDesde)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Saldo,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>= @lcFechaDesde AND Renglones_Comprobantes.Fec_Ini<= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Debe,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>= @lcFechaDesde AND Renglones_Comprobantes.Fec_Ini<= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Haber,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>= @lcFechaDesde AND Renglones_Comprobantes.Fec_Ini<= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Monto")
            loComandoSeleccionar.AppendLine("INTO	#tmpMovimientos")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables AS CC")
            loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ")
            loComandoSeleccionar.AppendLine("		INNER JOIN Comprobantes")
            loComandoSeleccionar.AppendLine("			ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine("				AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("				AND Comprobantes.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			)")
            loComandoSeleccionar.AppendLine("		ON CC.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")
            loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.Fec_Ini <= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("		LEFT JOIN Auxiliares ON Renglones_Comprobantes.Cod_Aux = Auxiliares.Cod_Aux")
            loComandoSeleccionar.AppendLine("WHERE	CC.Movimiento=1")
            loComandoSeleccionar.AppendLine("   AND CC.Categoria IN ('Activos', 'Pasivos', 'Capital')")
            loComandoSeleccionar.AppendLine("	AND	CC.Cod_Cue						BETWEEN @lcCuentaDesde	AND	@lcCuentaHasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Cod_Aux	BETWEEN @lcAuxDesde	AND	@lcAuxHasta")
            loComandoSeleccionar.AppendLine("GROUP BY CC.Cod_Cue, Renglones_Comprobantes.Cod_Aux, Auxiliares.Nom_Aux")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Cuentas_Contables.Cod_Cue							AS Cod_Cue,")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Nom_Cue							AS Nom_Cue,")
            If lnNivelMax = 5 Then
                loComandoSeleccionar.AppendLine("			#tmpMovimientos.Auxiliar							AS Auxiliar,")
                loComandoSeleccionar.AppendLine("           #tmpMovimientos.Nom_Auxiliar						AS Nom_Auxiliar,")
            End If
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_1, ")
            loComandoSeleccionar.AppendLine("			(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_1) AS Nom_Niv_1,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_2, ")
            loComandoSeleccionar.AppendLine("			(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_2) AS Nom_Niv_2,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_3, ")
            loComandoSeleccionar.AppendLine("			(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_3) AS Nom_Niv_3,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_4, ")
            loComandoSeleccionar.AppendLine("			(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_4) AS Nom_Niv_4,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_5, ")
            loComandoSeleccionar.AppendLine("			(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_5) AS Nom_Niv_5,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Saldo)							AS Saldo_Inicial,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Debe)							AS Debe,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Haber)							AS Haber,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto + #tmpMovimientos.Saldo)	AS Saldo_Actual")
            loComandoSeleccionar.AppendLine("INTO #tmpBalance")
            loComandoSeleccionar.AppendLine("FROM		#tmpMovimientos")
            loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Contables")
            loComandoSeleccionar.AppendLine("		ON	Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Cue")
            loComandoSeleccionar.AppendLine("GROUP BY	Cuentas_Contables.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Nom_Cue, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Movimiento,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_1, ")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_2, ")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_3,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_4, ")
            If lnNivelMax = 5 Then
                loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_5, ")
                loComandoSeleccionar.AppendLine("			#tmpMovimientos.Auxiliar,")
                loComandoSeleccionar.AppendLine("           #tmpMovimientos.Nom_Auxiliar")
            Else
                loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_5 ")
            End If
            If Not (llSoloMovimientos) Then
                loComandoSeleccionar.AppendLine("HAVING	ABS(SUM(#tmpMovimientos.Monto + #tmpMovimientos.Saldo)) > 0")
            End If
            loComandoSeleccionar.AppendLine("ORDER BY	Cod_Cue")
            If lcMostrarAux <> "Todos" And lnNivelMax = 5 Then
                loComandoSeleccionar.AppendLine("UPDATE #tmpBalance SET Auxiliar = '', Nom_Auxiliar = '' WHERE Cod_Cue <> @lcMostrarAux")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("SELECT Cod_Cue, Nom_Cue, SUM(Saldo_Inicial) AS Saldo_Inicial,")
                loComandoSeleccionar.AppendLine("       SUM(Debe) AS Debe, SUM(Haber) AS Haber, SUM(Saldo_Actual) AS Saldo_Actual")
                loComandoSeleccionar.AppendLine("INTO #tmpUpdate")
                loComandoSeleccionar.AppendLine("FROM #tmpBalance")
                loComandoSeleccionar.AppendLine("WHERE Cod_Cue <> @lcMostrarAux")
                loComandoSeleccionar.AppendLine("GROUP BY Cod_Cue, Nom_Cue")
                loComandoSeleccionar.AppendLine("ORDER BY Cod_Cue")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("UPDATE #tmpBalance SET Saldo_Inicial = (SELECT Saldo_Inicial FROM #tmpUpdate WHERE #tmpUpdate.Cod_Cue = #tmpBalance.Cod_Cue) ")
                loComandoSeleccionar.AppendLine("WHERE Cod_Cue <> @lcMostrarAux")
                loComandoSeleccionar.AppendLine("UPDATE #tmpBalance SET Debe = (SELECT Debe FROM #tmpUpdate WHERE #tmpUpdate.Cod_Cue = #tmpBalance.Cod_Cue) ")
                loComandoSeleccionar.AppendLine("WHERE Cod_Cue <> @lcMostrarAux")
                loComandoSeleccionar.AppendLine("UPDATE #tmpBalance SET Haber = (SELECT Haber FROM #tmpUpdate WHERE #tmpUpdate.Cod_Cue = #tmpBalance.Cod_Cue) ")
                loComandoSeleccionar.AppendLine("WHERE Cod_Cue <> @lcMostrarAux")
                loComandoSeleccionar.AppendLine("UPDATE #tmpBalance SET Saldo_Actual = (SELECT Saldo_Actual FROM #tmpUpdate WHERE #tmpUpdate.Cod_Cue = #tmpBalance.Cod_Cue) ")
                loComandoSeleccionar.AppendLine("WHERE Cod_Cue <> @lcMostrarAux ")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("DROP TABLE #tmpUpdate")
            End If
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("SELECT SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < @lcFechaDesde)")
            'loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            'loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            'loComandoSeleccionar.AppendLine("			END")
            'loComandoSeleccionar.AppendLine("		) ")
            'loComandoSeleccionar.AppendLine("		+")
            'loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>= @lcFechaDesde AND Renglones_Comprobantes.Fec_Ini<= @lcFechaHasta)")
            'loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            'loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            'loComandoSeleccionar.AppendLine("			END")
            'loComandoSeleccionar.AppendLine("		) AS Total")
            loComandoSeleccionar.AppendLine("SELECT SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < @lcFechaDesde)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Saldo,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>= @lcFechaDesde AND Renglones_Comprobantes.Fec_Ini<= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Debe,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>= @lcFechaDesde AND Renglones_Comprobantes.Fec_Ini<= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Haber")
            loComandoSeleccionar.AppendLine("INTO	#tmpEjercicio")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ON Cuentas_Contables.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")
            'loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.Fec_Ini BETWEEN @lcFechaDesde AND  @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.Fec_Ini <= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("WHERE	Cuentas_Contables.Movimiento=1")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Contables.Categoria NOT IN ('Activos', 'Pasivos', 'Capital')")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < @lcFechaDesde)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) ")
            loComandoSeleccionar.AppendLine("		+")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>= @lcFechaDesde AND Renglones_Comprobantes.Fec_Ini<= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Total")
            loComandoSeleccionar.AppendLine("INTO	#tmpPasivo")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ON Cuentas_Contables.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")
            loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.Fec_Ini <= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("WHERE	Cuentas_Contables.Movimiento=1")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Contables.Categoria = 'Pasivos'")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < @lcFechaDesde)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) ")
            loComandoSeleccionar.AppendLine("		+")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>= @lcFechaDesde AND Renglones_Comprobantes.Fec_Ini<= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Total")
            loComandoSeleccionar.AppendLine("INTO	#tmpCapital")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ON Cuentas_Contables.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")
            loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.Fec_Ini <= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("WHERE	Cuentas_Contables.Movimiento=1")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Contables.Categoria = 'Capital'")
            loComandoSeleccionar.AppendLine("")
            If lnNivelMax = 5 Then
                loComandoSeleccionar.AppendLine("SELECT DISTINCT *, @lcFechaHasta AS Hasta,")
                'loComandoSeleccionar.AppendLine("	(SELECT Total FROM #tmpEjercicio) AS Resultado_Ejercicio, ")
                loComandoSeleccionar.AppendLine("   (SELECT Saldo + (Debe - Haber) FROM #tmpEjercicio) AS Resultado_Ejercicio,")
                loComandoSeleccionar.AppendLine("	(SELECT Total FROM #tmpPasivo) AS Pasivo,  ")
                loComandoSeleccionar.AppendLine("	(SELECT Total FROM #tmpCapital) AS Capital  ")
                loComandoSeleccionar.AppendLine("FROM #tmpBalance ")
                loComandoSeleccionar.AppendLine("ORDER BY Cod_Cue")
            Else
                loComandoSeleccionar.AppendLine("SELECT DISTINCT *, '' AS Auxiliar, '' AS Nom_Auxiliar, @lcFechaHasta AS Hasta,")
                loComandoSeleccionar.AppendLine("	(SELECT Saldo + (Debe - Haber) FROM #tmpEjercicio) AS Resultado_Ejercicio, ")
                loComandoSeleccionar.AppendLine("	(SELECT Total FROM #tmpPasivo) AS Pasivo,  ")
                loComandoSeleccionar.AppendLine("	(SELECT Total FROM #tmpCapital) AS Capital  ")
                loComandoSeleccionar.AppendLine("FROM #tmpBalance ")
                loComandoSeleccionar.AppendLine("ORDER BY Cod_Cue")
            End If
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpBalance")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpEjercicio")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpPasivo")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCapital")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes", 180)

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

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rBalance_General", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rBalance_General.ReportSource = loObjetoReporte

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

