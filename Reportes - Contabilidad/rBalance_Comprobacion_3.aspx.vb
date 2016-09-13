'-------------------------------------------------------------------------------------------'
' Inicio del codigo																			'
'-------------------------------------------------------------------------------------------'
' Importando librerias																		'
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rBalance_Comprobacion_3"													'
'-------------------------------------------------------------------------------------------'
Partial Class rBalance_Comprobacion_3

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
			Dim lnTamaño1 AS Integer = CInt(goOpciones.mObtener("ANCNIV1", "")) 
			Dim lnTamaño2 AS Integer = CInt(goOpciones.mObtener("ANCNIV2", "")) 
			Dim lnTamaño3 AS Integer = CInt(goOpciones.mObtener("ANCNIV3", "")) 
			Dim lnTamaño4 AS Integer = CInt(goOpciones.mObtener("ANCNIV4", "")) 
			Dim lnTamaño5 AS Integer = CInt(goOpciones.mObtener("ANCNIV5", "")) 
			Dim lnTamaño6 AS Integer = CInt(goOpciones.mObtener("ANCNIV6", "")) 
			Dim lnTamaño7 AS Integer = 30 'CInt(goOpciones.mObtener("ANCNIV1", "")) 
			Dim lnCantidad AS Integer = CInt(goOpciones.mObtener("CANNIVCUE", "")) 
			Dim lcSeparador AS String = CStr(goOpciones.mObtener("CARSEPCUE", "")).Trim()
			
			lcSeparador = goServicios.mObtenerCampoFormatoSQL(lcSeparador)

			If lnCantidad<=0 Then lnCantidad = 1
			If (lnCantidad >= 1) And (lnTamaño1 <= 0)  Then 
				lnTamaño1 = 1 
			ElseIf (lnCantidad > 6) Then 
				lnCantidad = 6 
			End If
			
			If (lnCantidad >= 2) And (lnTamaño2 <= 0)  Then 
				lnTamaño2 = 1 
			ElseIf (lnCantidad < 2) Then 
				lnTamaño2 = 30
			End If
			lnTamaño2 += lnTamaño1
			
			If (lnCantidad >= 3) And (lnTamaño3 <= 0)  Then 
				lnTamaño3 = 1 
			ElseIf (lnCantidad < 3) Then 
				lnTamaño3 = 30
			End If
			lnTamaño3 += lnTamaño2
			
			If (lnCantidad >= 4) And (lnTamaño4 <= 0)  Then 
				lnTamaño4 = 1 
			ElseIf (lnCantidad < 4) Then 
				lnTamaño4 = 30
			End If
			lnTamaño4 += lnTamaño3
			
			If (lnCantidad >= 5) And (lnTamaño5 <= 0)  Then 
				lnTamaño5 = 1 
			ElseIf (lnCantidad < 5) Then 
				lnTamaño5 = 30
			End If
			lnTamaño5 += lnTamaño4
			
			If (lnCantidad >= 6) And (lnTamaño6 <= 0)  Then 
				lnTamaño6 = 1 
			ElseIf (lnCantidad < 6) Then 
				lnTamaño6 = 30
			End If
			lnTamaño6 += lnTamaño5
			lnTamaño7 += lnTamaño6
			

			loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Desde DATETIME	")		'Fecha inicio
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Hasta DATETIME	")		'Fecha Fin
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Desde VARCHAR(100)	")	'Cuenta Contable
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Hasta VARCHAR(100)	")	'
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Desde VARCHAR(100)	")	'Centro costo
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Hasta VARCHAR(100)	")	'
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Desde VARCHAR(100)	")	'Cta. gasto
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Hasta VARCHAR(100)	")	'
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Desde VARCHAR(100)	")	'Auxiliares 
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Hasta VARCHAR(100)	")	'
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro6Hasta VARCHAR(100)	")	'moneda
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro6Desde VARCHAR(100)	")	'
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro7Desde VARCHAR(100)	")	'Tipo
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro8Desde VARCHAR(100)	")	'Con Movimientos
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro9Desde VARCHAR(100)	")	'Nivel
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
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("--*********************************************************")
			loComandoSeleccionar.AppendLine("-- Prepara un listado de las cuentas contables a incluir  *")
			loComandoSeleccionar.AppendLine("--*********************************************************")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10);")
			loComandoSeleccionar.AppendLine("SET		@lnCero = 0;")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño1 INT; SET @lnTamaño1= " & lnTamaño1.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño2 INT; SET @lnTamaño2= " & lnTamaño2.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño3 INT; SET @lnTamaño3= " & lnTamaño3.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño4 INT; SET @lnTamaño4= " & lnTamaño4.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño5 INT; SET @lnTamaño5= " & lnTamaño5.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño6 INT; SET @lnTamaño6= " & lnTamaño6.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño7 INT; SET @lnTamaño7= " & lnTamaño7.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamañoMax INT; SET @lnTamañoMax = @lnTamaño7;")
			loComandoSeleccionar.AppendLine("DECLARE @llUsarTasa BIT; SET @llUsarTasa = 0;")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("IF(@lcParametro9Desde > " & lnCantidad & ") SET @lcParametro9Desde = " & lnCantidad & " ")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("IF (@lcParametro9Desde = 1)   ")
			loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño1")
			loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde = 2)")
			loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño2")
			loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde = 3)")
			loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño3 ")
			loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde = 4)")
			loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño4 ")
			loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde = 5)")
			loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño5;")
			loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde = 6)")
			loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño6;")
			loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde >= 7)")
			loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño7;")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lnMaximaLongitud AS INT;")
			loComandoSeleccionar.AppendLine("SET @lnMaximaLongitud = (SELECT MAX(LEN(SUBSTRING(Cuentas_Contables.Cod_Cue,1,@lnTamañoMax))) FROM Cuentas_Contables)")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	Cuentas_Contables.Cod_Cue	AS Cod_Cue,")
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < @lcParametro1Desde)")
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.mon_deb - Renglones_Comprobantes.mon_hab, @lnCero)")
			loComandoSeleccionar.AppendLine("					ELSE @lnCero")
			loComandoSeleccionar.AppendLine("			END")
			loComandoSeleccionar.AppendLine("		) AS saldo,")
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=@lcParametro1Desde and Renglones_Comprobantes.Fec_Ini<=@lcParametro1Hasta)")
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.mon_deb, @lnCero)")
			loComandoSeleccionar.AppendLine("					ELSE @lnCero")
			loComandoSeleccionar.AppendLine("			END")
			loComandoSeleccionar.AppendLine("		) AS Debe,")
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=@lcParametro1Desde and Renglones_Comprobantes.Fec_Ini<=@lcParametro1Hasta)")
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Hab, @lnCero)")
			loComandoSeleccionar.AppendLine("					ELSE @lnCero")
			loComandoSeleccionar.AppendLine("			END")
			loComandoSeleccionar.AppendLine("		) AS Haber,")
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=@lcParametro1Desde and Renglones_Comprobantes.Fec_Ini<=@lcParametro1Hasta)")
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.mon_deb - Renglones_Comprobantes.mon_hab, @lnCero)")
			loComandoSeleccionar.AppendLine("					ELSE @lnCero")
			loComandoSeleccionar.AppendLine("			END")
			loComandoSeleccionar.AppendLine("		) AS Monto		")
			loComandoSeleccionar.AppendLine("INTO	#tmpValores")
			loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables ")
			loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ")
			loComandoSeleccionar.AppendLine("		inner JOIN Comprobantes")
			loComandoSeleccionar.AppendLine("			ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
			loComandoSeleccionar.AppendLine("				AND Renglones_Comprobantes.documento = Comprobantes.documento")
			loComandoSeleccionar.AppendLine("				AND Comprobantes.Tipo =@lcParametro7Desde and Comprobantes.Status<> 'anulado'")
			loComandoSeleccionar.AppendLine("			)")
			loComandoSeleccionar.AppendLine("		ON Cuentas_Contables.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")
			loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.fec_ini<=@lcParametro1Hasta)")
			loComandoSeleccionar.AppendLine("WHERE	Cuentas_Contables.Movimiento=1")
			loComandoSeleccionar.AppendLine("		AND Cuentas_Contables.cod_cue		BETWEEN " & lcParametro1Desde & "	AND	" & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.cod_cen	BETWEEN " & lcParametro2Desde & "	AND	" & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.cod_gas	BETWEEN " & lcParametro3Desde & "	AND	" & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.cod_aux	BETWEEN " & lcParametro4Desde & "	AND	" & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.cod_mon	BETWEEN " & lcParametro5Desde & "	AND	" & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine("GROUP BY Cuentas_Contables.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	CC.Cod_Cue																					AS Cod_Cue, ")
			loComandoSeleccionar.AppendLine("		CC.Nom_Cue																					AS Nom_cue,")
			'loComandoSeleccionar.AppendLine("		(CASE WHEN (LEN(CC.Cod_Cue) = @lnMaximaLongitud) THEN 1 ELSE 0 END)							AS Movimiento,")
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN @lnTamaño1 <= @lnTamañoMax THEN @lnTamaño1 ELSE 0 END)	AS Nivel_1,")
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN @lnTamaño2 <= @lnTamañoMax THEN @lnTamaño2 ELSE 0 END)	AS Nivel_2,")
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN @lnTamaño3 <= @lnTamañoMax THEN @lnTamaño3 ELSE 0 END)	AS Nivel_3,")
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN @lnTamaño4 <= @lnTamañoMax THEN @lnTamaño4 ELSE 0 END)	AS Nivel_4,")
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN @lnTamaño5 <= @lnTamañoMax THEN @lnTamaño5 ELSE 0 END)	AS Nivel_5,")
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN @lnTamaño6 <= @lnTamañoMax THEN @lnTamaño6 ELSE 0 END)	AS Nivel_6,")
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN @lnTamaño7 <= @lnTamañoMax THEN @lnTamaño7 ELSE 0 END)	AS Nivel_7, ")
			loComandoSeleccionar.AppendLine("		CASE CC.Cod_Cue")
			loComandoSeleccionar.AppendLine("			WHEN  SUBSTRING(CC.Cod_Cue, 1, @lnTamaño1) THEN 1 ")
			loComandoSeleccionar.AppendLine("			WHEN  SUBSTRING(CC.Cod_Cue, 1, @lnTamaño2) THEN 2 ")
			loComandoSeleccionar.AppendLine("			WHEN  SUBSTRING(CC.Cod_Cue, 1, @lnTamaño3) THEN 3 ")
			loComandoSeleccionar.AppendLine("			WHEN  SUBSTRING(CC.Cod_Cue, 1, @lnTamaño4) THEN 4 ")
			loComandoSeleccionar.AppendLine("			WHEN  SUBSTRING(CC.Cod_Cue, 1, @lnTamaño5) THEN 5 ")
			loComandoSeleccionar.AppendLine("			WHEN  SUBSTRING(CC.Cod_Cue, 1, @lnTamaño6) THEN 6 ")
			loComandoSeleccionar.AppendLine("			WHEN  SUBSTRING(CC.Cod_Cue, 1, @lnTamaño7) THEN 7 ")
			loComandoSeleccionar.AppendLine("			ELSE 10")
			loComandoSeleccionar.AppendLine("		END																							AS Nivel,")
			loComandoSeleccionar.AppendLine("		@lnMaximaLongitud																			AS Maximo,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpValores.Saldo, @lnCero)															AS Saldo_Inicial,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpValores.Debe, @lnCero)															AS Debe,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpValores.Haber, @lnCero)															AS Haber,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpValores.Monto, @lnCero)															AS Monto,")
			loComandoSeleccionar.AppendLine("		ISNULL(#tmpValores.Saldo + #tmpValores.Monto, @lnCero)										AS Saldo_Actual")
			loComandoSeleccionar.AppendLine("INTO	#tmpParcial")
			loComandoSeleccionar.AppendLine("FROM	#tmpValores")
			loComandoSeleccionar.AppendLine("	RIGHT JOIN Cuentas_Contables As CC ")
			loComandoSeleccionar.AppendLine("		ON #tmpValores.Cod_Cue = CC.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpValores")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT		A.Cod_Cue, A.Nom_Cue, ")
			loComandoSeleccionar.AppendLine("			CAST((CASE WHEN A.Nivel=@lcParametro9Desde THEN 1 ELSE 0 END) AS BIT) AS  Movimiento,")
			loComandoSeleccionar.AppendLine("			A.Nivel_1, A.Nivel_2, A.Nivel_3, A.Nivel_4, ")
			loComandoSeleccionar.AppendLine("			A.Nivel_5, A.Nivel_6, A.Nivel_7, A.Nivel, A.Maximo, ")
			loComandoSeleccionar.AppendLine("			SUM(ISNULL(B.Saldo_Inicial, A.Saldo_Inicial))	AS Saldo_Inicial, ")
			loComandoSeleccionar.AppendLine("			SUM(ISNULL(B.Debe, A.Debe))						AS Debe, ")
			loComandoSeleccionar.AppendLine("			SUM(ISNULL(B.Haber, A.Haber))					AS Haber, ")
			loComandoSeleccionar.AppendLine("			SUM(ISNULL(B.Monto, A.Monto))					AS Monto, ")
			loComandoSeleccionar.AppendLine("			SUM(ISNULL(B.Saldo_Actual, A.Saldo_Actual))		AS Saldo_Actual  ")
			loComandoSeleccionar.AppendLine("INTO		#tmpListadoAgrupado")
			loComandoSeleccionar.AppendLine("FROM		#tmpParcial AS A")
			loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpParcial AS B ON B.Cod_Cue LIKE (RTRIM(A.Cod_Cue)+'%')")
			loComandoSeleccionar.AppendLine("		AND	(CASE WHEN A.Nivel=@lcParametro9Desde THEN 1 ELSE 0 END) > 0")
			loComandoSeleccionar.AppendLine("		AND LEN(B.Cod_Cue) >= A.Maximo")
			loComandoSeleccionar.AppendLine("WHERE		A.Nivel<=@lcParametro9Desde")
			loComandoSeleccionar.AppendLine("GROUP BY	A.Cod_Cue, A.Nom_Cue,  ")
			loComandoSeleccionar.AppendLine("			A.Nivel_1, A.Nivel_2, A.Nivel_3, A.Nivel_4, ")
			loComandoSeleccionar.AppendLine("			A.Nivel_5, A.Nivel_6, A.Nivel_7, A.Nivel, A.Maximo")
			If llSoloMovimientos Then 
				loComandoSeleccionar.AppendLine("HAVING		ABS(SUM(ISNULL(B.Saldo_Inicial, @lnCero))) ")
				loComandoSeleccionar.AppendLine("		  + ABS(SUM(ISNULL(B.Debe, @lnCero))) ")
				loComandoSeleccionar.AppendLine("		  + ABS(SUM(ISNULL(B.Haber, @lnCero))) > 0")
				loComandoSeleccionar.AppendLine("		OR A.Nivel<@lcParametro9Desde ")
				loComandoSeleccionar.AppendLine("		OR A.Nivel<@lcParametro9Desde ")
			End If
			loComandoSeleccionar.AppendLine("ORDER BY	A.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpParcial")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")								  
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DELETE FROM #tmpListadoAgrupado ")
			loComandoSeleccionar.AppendLine("FROM		(	SELECT	A.Cod_cue")
			loComandoSeleccionar.AppendLine("				FROM	#tmpListadoAgrupado As A")
			loComandoSeleccionar.AppendLine("				WHERE	(SELECT COUNT(1) FROM #tmpListadoAgrupado AS B WHERE B.Nivel_7 = RTRIM(A.Cod_cue)+" & lcSeparador & ") <= 1")
			loComandoSeleccionar.AppendLine("					AND A.Nivel = 7 AND A.Movimiento = 0")
			loComandoSeleccionar.AppendLine("			) AS Borrar")
			loComandoSeleccionar.AppendLine("WHERE Borrar.Cod_Cue = #tmpListadoAgrupado.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DELETE FROM #tmpListadoAgrupado ")
			loComandoSeleccionar.AppendLine("FROM		(	SELECT	A.Cod_cue")
			loComandoSeleccionar.AppendLine("				FROM	#tmpListadoAgrupado As A")
			loComandoSeleccionar.AppendLine("				WHERE	(SELECT COUNT(1) FROM #tmpListadoAgrupado AS B WHERE B.Nivel_6 = RTRIM(A.Cod_cue)+" & lcSeparador & ") <= 1")
			loComandoSeleccionar.AppendLine("					AND A.Nivel = 6 AND A.Movimiento = 0")
			loComandoSeleccionar.AppendLine("			) AS Borrar")
			loComandoSeleccionar.AppendLine("WHERE Borrar.Cod_Cue = #tmpListadoAgrupado.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DELETE FROM #tmpListadoAgrupado ")
			loComandoSeleccionar.AppendLine("FROM		(	SELECT	A.Cod_cue")
			loComandoSeleccionar.AppendLine("				FROM	#tmpListadoAgrupado As A")
			loComandoSeleccionar.AppendLine("				WHERE	(SELECT COUNT(1) FROM #tmpListadoAgrupado AS B WHERE B.Nivel_5 = RTRIM(A.Cod_cue)+" & lcSeparador & ") <= 1")
			loComandoSeleccionar.AppendLine("					AND A.Nivel = 5 AND A.Movimiento = 0")
			loComandoSeleccionar.AppendLine("			) AS Borrar")
			loComandoSeleccionar.AppendLine("WHERE Borrar.Cod_Cue = #tmpListadoAgrupado.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DELETE FROM #tmpListadoAgrupado ")
			loComandoSeleccionar.AppendLine("FROM		(	SELECT	A.Cod_cue")
			loComandoSeleccionar.AppendLine("				FROM	#tmpListadoAgrupado As A")
			loComandoSeleccionar.AppendLine("				WHERE	(SELECT COUNT(1) FROM #tmpListadoAgrupado AS B WHERE B.Nivel_4 = RTRIM(A.Cod_cue)+" & lcSeparador & ") <= 1")
			loComandoSeleccionar.AppendLine("					AND A.Nivel = 4 AND A.Movimiento = 0")
			loComandoSeleccionar.AppendLine("			) AS Borrar")
			loComandoSeleccionar.AppendLine("WHERE Borrar.Cod_Cue = #tmpListadoAgrupado.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DELETE FROM #tmpListadoAgrupado ")
			loComandoSeleccionar.AppendLine("FROM		(	SELECT	A.Cod_cue")
			loComandoSeleccionar.AppendLine("				FROM	#tmpListadoAgrupado As A")
			loComandoSeleccionar.AppendLine("				WHERE	(SELECT COUNT(1) FROM #tmpListadoAgrupado AS B WHERE B.Nivel_3 = RTRIM(A.Cod_cue)+" & lcSeparador & ") <= 1")
			loComandoSeleccionar.AppendLine("					AND A.Nivel = 3 AND A.Movimiento = 0")
			loComandoSeleccionar.AppendLine("			) AS Borrar")
			loComandoSeleccionar.AppendLine("WHERE Borrar.Cod_Cue = #tmpListadoAgrupado.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DELETE FROM #tmpListadoAgrupado ")
			loComandoSeleccionar.AppendLine("FROM		(	SELECT	A.Cod_cue")
			loComandoSeleccionar.AppendLine("				FROM	#tmpListadoAgrupado As A")
			loComandoSeleccionar.AppendLine("				WHERE	(SELECT COUNT(1) FROM #tmpListadoAgrupado AS B WHERE B.Nivel_2 = RTRIM(A.Cod_cue)+" & lcSeparador & ") <= 1")
			loComandoSeleccionar.AppendLine("					AND A.Nivel = 2 AND A.Movimiento = 0")
			loComandoSeleccionar.AppendLine("			) AS Borrar")
			loComandoSeleccionar.AppendLine("WHERE Borrar.Cod_Cue = #tmpListadoAgrupado.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DELETE FROM #tmpListadoAgrupado ")
			loComandoSeleccionar.AppendLine("FROM		(	SELECT	A.Cod_cue")
			loComandoSeleccionar.AppendLine("				FROM	#tmpListadoAgrupado As A")
			loComandoSeleccionar.AppendLine("				WHERE	(SELECT COUNT(1) FROM #tmpListadoAgrupado AS B WHERE B.Nivel_1 = RTRIM(A.Cod_cue)+" & lcSeparador & ") <= 1")
			loComandoSeleccionar.AppendLine("					AND A.Nivel = 1 AND A.Movimiento = 0")
			loComandoSeleccionar.AppendLine("			) AS Borrar")
			loComandoSeleccionar.AppendLine("WHERE Borrar.Cod_Cue = #tmpListadoAgrupado.Cod_Cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT * FROM #tmpListadoAgrupado")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpListadoAgrupado")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes", 900)

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

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rBalance_Comprobacion_3", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrBalance_Comprobacion_3.ReportSource = loObjetoReporte

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
' RJG: 26/10/11: Codigo inicial, a partir de rBalance_Comprobacion_2 (mejora de Velocidad).	'
'-------------------------------------------------------------------------------------------' 
' RJG: 09/02/12: Cambiado el filtro de comprobantes "='Pendiente'" por "<>'Anulado'".		'
'-------------------------------------------------------------------------------------------' 
' RJG: 10/02/12: Corrección en etiquetas de cierre de subtotales: no se actualizaban al		'
'				 pasar de un nivel a otro.													'
'-------------------------------------------------------------------------------------------' 
