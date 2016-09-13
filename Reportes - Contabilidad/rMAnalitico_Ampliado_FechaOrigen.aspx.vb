'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMAnalitico_Ampliado_FechaOrigen"
'-------------------------------------------------------------------------------------------'
Partial Class rMAnalitico_Ampliado_FechaOrigen

	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
	
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try

			Dim lcFechaDesde 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcFechaHasta 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcCuentaContableDesde	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcCuentaContableHasta	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
			Dim lcCentroCostoDesde		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcCentroCostoHasta		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcCuentaGastoDesde		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcCuentaGastoHasta		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			Dim lcAuxiliarDesde			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcAuxiliarHasta			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcMonedaDesde 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcMonedaHasta 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcDocumentoDesde		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcDocumentoHasta		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			'Dim lcSoloMovimientos		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))			
					
 			Dim llSoloMovimientos As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(7)).Trim().ToUpper().Equals("SI")
		
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden


			Dim lnTamaño1 As Integer = CInt(goOpciones.mObtener("ANCNIV1", "")) 
			Dim lnTamaño2 As Integer = CInt(goOpciones.mObtener("ANCNIV2", "")) 
			Dim lnTamaño3 As Integer = CInt(goOpciones.mObtener("ANCNIV3", "")) 
			Dim lnTamaño4 As Integer = CInt(goOpciones.mObtener("ANCNIV4", "")) 
			Dim lnTamaño5 As Integer = CInt(goOpciones.mObtener("ANCNIV5", "")) 
			Dim lnTamaño6 As Integer = CInt(goOpciones.mObtener("ANCNIV6", ""))
			
			Dim lnCantidad As Integer = CInt(goOpciones.mObtener("CANNIVCUE", ""))				'Niveles usados por la empresa
			Dim lnNivelMax As Integer = 7														'Nivel máximo a mostrar	en reporte: Todos lso niveles
			Dim lcSeparador As String = CStr(goOpciones.mObtener("CARSEPCUE", "")).Trim()		'Separador de Niveles de la empresa

			Dim lnLongMax As Integer = 0

		'Establece el nivel máximo a mostrar
			If lnCantidad<=0 Then
				 lnCantidad = 1
			ElseIf (lnCantidad > 6) Then 
				lnCantidad = 6 
			End If
			lnCantidad = Math.Min(lnCantidad, lnNivelMax)
			
		'Establece las longitudes de todos los niveles
			
			If (lnCantidad < 1) Then 
				lnTamaño1 = 0
			ElseIf (lnTamaño1 <= 0)  Then 
				lnTamaño1 = Math.Max(lnTamaño1, 1) 
			End If
			
			If (lnCantidad < 2) Then 
				lnTamaño2 = 0
			ElseIf (lnTamaño2 <= 0)  Then 
				lnTamaño2 = Math.Max(lnTamaño2, 1) 
			End If
			lnTamaño2 += lnTamaño1
			
			If (lnCantidad < 3) Then 
				lnTamaño3 = 0
			ElseIf (lnTamaño3 <= 0)  Then 
				lnTamaño3 = Math.Max(lnTamaño3, 1) 
			End If
			lnTamaño3 += lnTamaño2

			If (lnCantidad < 4) Then 
				lnTamaño4 = 0
			ElseIf (lnTamaño4 <= 0)  Then 
				lnTamaño4 = Math.Max(lnTamaño4, 1) 
			End If
			lnTamaño4 += lnTamaño3

			If (lnCantidad < 5) Then 
				lnTamaño5 = 0
			ElseIf (lnTamaño5 <= 0)  Then 
				lnTamaño5 = Math.Max(lnTamaño5, 1) 
			End If
			lnTamaño5 += lnTamaño4

			If (lnCantidad < 6) Then 
				lnTamaño6 = 0
			ElseIf (lnTamaño6 <= 0)  Then 
				lnTamaño6 = Math.Max(lnTamaño6, 1) 
			End If
			lnTamaño6 += lnTamaño5

			Select Case lnNivelMax
				Case 1
					lnLongMax = lnTamaño1 
				Case 2
					lnLongMax = lnTamaño2 
				Case 3
					lnLongMax = lnTamaño3 
				Case 4
					lnLongMax = lnTamaño4 
				Case 5
					lnLongMax = lnTamaño5 
				Case 6
					lnLongMax = lnTamaño6 
				Case 7	'Mostrar Todos
					lnLongMax = lnTamaño6 
			End Select


			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE	@lnCero DECIMAL(28, 10);")
			loComandoSeleccionar.AppendLine("SET		@lnCero = 0;")
			loComandoSeleccionar.AppendLine("DECLARE @llUsarTasa BIT; SET @llUsarTasa = 0;")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño1 INT; SET @lnTamaño1= " & lnTamaño1.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño2 INT; SET @lnTamaño2= " & lnTamaño2.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño3 INT; SET @lnTamaño3= " & lnTamaño3.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño4 INT; SET @lnTamaño4= " & lnTamaño4.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño5 INT; SET @lnTamaño5= " & lnTamaño5.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño6 INT; SET @lnTamaño6= " & lnTamaño6.ToString() & ";")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lnNivelMax AS INT;")
			loComandoSeleccionar.AppendLine("SET @lnNivelMax = " & lnNivelMax.ToString())
			loComandoSeleccionar.AppendLine("DECLARE @lnLongMaxima AS INT;")
			loComandoSeleccionar.AppendLine("SET @lnLongMaxima = " & lnLongMax.ToString())
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("SELECT	ROW_NUMBER() ") 
			loComandoSeleccionar.AppendLine("			OVER (PARTITION BY Renglones_Comprobantes.Cod_Cue ") 
			loComandoSeleccionar.AppendLine("				ORDER BY Renglones_Comprobantes.Documento, Renglones_Comprobantes.Renglon)										AS Posicion, ") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño1 AND @lnNivelMax > 1) THEN @lnTamaño1 ELSE 0 END)	AS Cod_Niv_1,") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño2 AND @lnNivelMax > 2) THEN @lnTamaño2 ELSE 0 END)	AS Cod_Niv_2,") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño3 AND @lnNivelMax > 3) THEN @lnTamaño3 ELSE 0 END)	AS Cod_Niv_3,") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño4 AND @lnNivelMax > 4) THEN @lnTamaño4 ELSE 0 END)	AS Cod_Niv_4,") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño5 AND @lnNivelMax > 5) THEN @lnTamaño5 ELSE 0 END)	AS Cod_Niv_5,") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño6 AND @lnNivelMax > 6) THEN @lnTamaño6 ELSE 0 END)	AS Cod_Niv_6,") 
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Documento																						AS Documento, ")
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Renglon																							AS Renglon, ") 
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Referencia																						AS Referencia, ") 
			loComandoSeleccionar.AppendLine("		CAST(Renglones_Comprobantes.Comentario AS VARCHAR(MAX))																	AS Comentario, ") 
			loComandoSeleccionar.AppendLine("		CC.Cod_Cue																												AS Cod_Cue, ")
			loComandoSeleccionar.AppendLine("		CC.Nom_Cue																												AS Nom_Cue, ")
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Cen																							AS Cod_Cen, ") 
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Gas																							AS Cod_Gas, ") 
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Aux																							AS Cod_Aux, ") 
			loComandoSeleccionar.AppendLine("		ISNULL(Auxiliares.Nom_Aux, '')																							AS Nom_Aux, ")
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Fec_Ori																							AS Fec_Ori, ") 
			loComandoSeleccionar.AppendLine("		ISNULL(Iniciales.Saldo_Inicial, @lnCero)																				AS Saldo_Inicial,") 
			loComandoSeleccionar.AppendLine("		(ISNULL(Renglones_Comprobantes.Mon_Deb, @lnCero)") 
			loComandoSeleccionar.AppendLine("				*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)")
			loComandoSeleccionar.AppendLine("			)																													AS Debe,") 
			loComandoSeleccionar.AppendLine("		(ISNULL(Renglones_Comprobantes.Mon_Hab, @lnCero)") 
			loComandoSeleccionar.AppendLine("				*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("			)																													AS Haber")
			loComandoSeleccionar.AppendLine("INTO 	#tmpMovimientos") 
			loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables AS CC")
			loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ") 
			loComandoSeleccionar.AppendLine("		INNER JOIN Comprobantes") 
			loComandoSeleccionar.AppendLine("			ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
			loComandoSeleccionar.AppendLine("				AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
			loComandoSeleccionar.AppendLine("				AND Comprobantes.Tipo = 'Diario' AND Comprobantes.Status <> 'Anulado'") 
			loComandoSeleccionar.AppendLine("				AND Comprobantes.Documento BETWEEN " & lcDocumentoDesde & " AND " & lcDocumentoHasta)
			loComandoSeleccionar.AppendLine("			)") 
			loComandoSeleccionar.AppendLine("		ON	CC.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")
			loComandoSeleccionar.AppendLine("		AND (Renglones_Comprobantes.Fec_Ori BETWEEN " & lcFechaDesde & " AND " & lcFechaHasta & ")")
			loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN (") 
			loComandoSeleccionar.AppendLine("			SELECT	RC.Cod_Cue						AS Cod_Cue,") 
			loComandoSeleccionar.AppendLine("					SUM(RC.Mon_Deb - RC.Mon_Hab)	AS Saldo_Inicial")
			loComandoSeleccionar.AppendLine("			FROM	Renglones_Comprobantes AS RC ") 
			loComandoSeleccionar.AppendLine("				INNER JOIN Comprobantes AS C") 
			loComandoSeleccionar.AppendLine("					ON (RC.Adicional = C.Adicional ") 
			loComandoSeleccionar.AppendLine("						AND RC.Documento = C.Documento")
			loComandoSeleccionar.AppendLine("						AND C.Tipo = 'Diario' AND C.Status <> 'Anulado'") 
			loComandoSeleccionar.AppendLine("						AND C.Documento BETWEEN " & lcDocumentoDesde & " AND " & lcDocumentoHasta & ")")
			loComandoSeleccionar.AppendLine("			WHERE	(RC.Fec_Ori < " & lcFechaDesde & " )") 
			If llSoloMovimientos Then
				loComandoSeleccionar.AppendLine("		AND RC.Cod_Cen	BETWEEN " & lcCentroCostoDesde		& "	AND	" & lcCentroCostoHasta)
				loComandoSeleccionar.AppendLine("		AND RC.Cod_Gas	BETWEEN " & lcCuentaGastoDesde		& "	AND	" & lcCuentaGastoHasta)
				loComandoSeleccionar.AppendLine("		AND RC.Cod_Aux	BETWEEN " & lcAuxiliarDesde			& "	AND	" & lcAuxiliarHasta)
				loComandoSeleccionar.AppendLine("		AND RC.Cod_Mon	BETWEEN " & lcMonedaDesde			& "	AND	" & lcMonedaHasta)
			Else
				loComandoSeleccionar.AppendLine("		AND ISNULL(RC.Cod_Cen, '')	BETWEEN " & lcCentroCostoDesde		& "	AND	" & lcCentroCostoHasta)
				loComandoSeleccionar.AppendLine("		AND ISNULL(RC.Cod_Gas, '')	BETWEEN " & lcCuentaGastoDesde		& "	AND	" & lcCuentaGastoHasta)
				loComandoSeleccionar.AppendLine("		AND ISNULL(RC.Cod_Aux, '')	BETWEEN " & lcAuxiliarDesde			& "	AND	" & lcAuxiliarHasta)
				loComandoSeleccionar.AppendLine("		AND ISNULL(RC.Cod_Mon, '')	BETWEEN " & lcMonedaDesde			& "	AND	" & lcMonedaHasta)
			End If
			loComandoSeleccionar.AppendLine("			GROUP BY RC.Cod_Cue")
			loComandoSeleccionar.AppendLine("			) AS Iniciales") 
			loComandoSeleccionar.AppendLine("		ON CC.Cod_Cue = Iniciales.Cod_Cue ") 
			loComandoSeleccionar.AppendLine("	LEFT JOIN Auxiliares")
			loComandoSeleccionar.AppendLine("       ON	Renglones_Comprobantes.Cod_Aux = Auxiliares.Cod_Aux") 
			loComandoSeleccionar.AppendLine("WHERE	CC.Movimiento = 1") 
			loComandoSeleccionar.AppendLine("		AND	CC.Cod_Cue						BETWEEN " & lcCuentaContableDesde	& "	AND	" & lcCuentaContableHasta)
			If llSoloMovimientos Then
				loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Cen	BETWEEN " & lcCentroCostoDesde	& "	AND	" & lcCentroCostoHasta)
				loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Gas	BETWEEN " & lcCuentaGastoDesde	& "	AND	" & lcCuentaGastoHasta)
				loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Aux	BETWEEN " & lcAuxiliarDesde		& "	AND	" & lcAuxiliarHasta)
				loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Mon	BETWEEN " & lcMonedaDesde		& "	AND	" & lcMonedaHasta)
			Else
				loComandoSeleccionar.AppendLine("		AND ISNULL(Renglones_Comprobantes.Cod_Cen, " & lcCentroCostoDesde	& ")	BETWEEN " & lcCentroCostoDesde	& "	AND	" & lcCentroCostoHasta)
				loComandoSeleccionar.AppendLine("		AND ISNULL(Renglones_Comprobantes.Cod_Gas, " & lcCuentaGastoDesde	& ")	BETWEEN " & lcCuentaGastoDesde	& "	AND	" & lcCuentaGastoHasta)
				loComandoSeleccionar.AppendLine("		AND ISNULL(Renglones_Comprobantes.Cod_Aux, " & lcAuxiliarDesde		& ")	BETWEEN " & lcAuxiliarDesde		& "	AND	" & lcAuxiliarHasta)
				loComandoSeleccionar.AppendLine("		AND ISNULL(Renglones_Comprobantes.Cod_Mon, " & lcMonedaDesde		& ")	BETWEEN " & lcMonedaDesde		& "	AND	" & lcMonedaHasta)
			End If
			loComandoSeleccionar.AppendLine("ORDER BY CC.Cod_Cue, Posicion")
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("CREATE Clustered INDEX PK_tmpMovimientos_Cuenta_Posicion ON #tmpMovimientos(Cod_Cue, Posicion)")
			
			If (lcSeparador <> "") Then
				Dim lcSeparadorSQL  As String = goServicios.mObtenerCampoFormatoSQL("%" & lcSeparador)
				loComandoSeleccionar.AppendLine("") 
				loComandoSeleccionar.AppendLine("UPDATE	#tmpMovimientos") 
				loComandoSeleccionar.AppendLine("SET		Cod_Cue = SUBSTRING(Cod_Cue, 1, LEN(Cod_Cue) - 1)") 
				loComandoSeleccionar.AppendLine("WHERE	ISNULL(Cod_Cue, '') LIKE " & lcSeparadorSQL ) 
				loComandoSeleccionar.AppendLine("") 
				loComandoSeleccionar.AppendLine("UPDATE	#tmpMovimientos") 
				loComandoSeleccionar.AppendLine("SET		Cod_Niv_1 = SUBSTRING(Cod_Niv_1, 1, LEN(Cod_Niv_1) - 1)") 
				loComandoSeleccionar.AppendLine("WHERE	ISNULL(Cod_Niv_1, '') LIKE " & lcSeparadorSQL )
				loComandoSeleccionar.AppendLine("") 
				loComandoSeleccionar.AppendLine("UPDATE	#tmpMovimientos									") 
				loComandoSeleccionar.AppendLine("SET		Cod_Niv_2 = SUBSTRING(Cod_Niv_2, 1, LEN(Cod_Niv_2) - 1)") 
				loComandoSeleccionar.AppendLine("WHERE	ISNULL(Cod_Niv_2, '') LIKE " & lcSeparadorSQL ) 
				loComandoSeleccionar.AppendLine("") 
				loComandoSeleccionar.AppendLine("UPDATE		#tmpMovimientos									") 
				loComandoSeleccionar.AppendLine("SET			Cod_Niv_3 = SUBSTRING(Cod_Niv_3, 1, LEN(Cod_Niv_3) - 1)") 
				loComandoSeleccionar.AppendLine("WHERE	ISNULL(Cod_Niv_3, '') LIKE " & lcSeparadorSQL ) 
				loComandoSeleccionar.AppendLine("") 
				loComandoSeleccionar.AppendLine("UPDATE		#tmpMovimientos									") 
				loComandoSeleccionar.AppendLine("SET			Cod_Niv_4 = SUBSTRING(Cod_Niv_4, 1, LEN(Cod_Niv_4) - 1)") 
				loComandoSeleccionar.AppendLine("WHERE	ISNULL(Cod_Niv_4, '') LIKE " & lcSeparadorSQL ) 
				loComandoSeleccionar.AppendLine("") 
				loComandoSeleccionar.AppendLine("UPDATE		#tmpMovimientos									") 
				loComandoSeleccionar.AppendLine("SET			Cod_Niv_5 = SUBSTRING(Cod_Niv_5, 1, LEN(Cod_Niv_5) - 1)") 
				loComandoSeleccionar.AppendLine("WHERE	ISNULL(Cod_Niv_5, '') LIKE " & lcSeparadorSQL ) 
				loComandoSeleccionar.AppendLine("") 
				loComandoSeleccionar.AppendLine("UPDATE		#tmpMovimientos									") 
				loComandoSeleccionar.AppendLine("SET			Cod_Niv_6 = SUBSTRING(Cod_Niv_6, 1, LEN(Cod_Niv_6) - 1)") 
				loComandoSeleccionar.AppendLine("WHERE	ISNULL(Cod_Niv_6, '') LIKE " & lcSeparadorSQL ) 
				loComandoSeleccionar.AppendLine("") 
			End If

			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	#tmpMovimientos.Posicion 										AS Posicion,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Niv_1										AS Cod_Niv_1,	") 
			loComandoSeleccionar.AppendLine("		(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables									") 
			loComandoSeleccionar.AppendLine("			WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_1)AS Nom_Niv_1,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Niv_2										AS Cod_Niv_2,	") 
			loComandoSeleccionar.AppendLine("		(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables									") 
			loComandoSeleccionar.AppendLine("			WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_2)AS Nom_Niv_2,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Niv_3										AS Cod_Niv_3,	") 
			loComandoSeleccionar.AppendLine("		(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables									") 
			loComandoSeleccionar.AppendLine("			WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_3)AS Nom_Niv_3,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Niv_4										AS Cod_Niv_4,	") 
			loComandoSeleccionar.AppendLine("		(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables									")
			loComandoSeleccionar.AppendLine("			WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_4)AS Nom_Niv_4,	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Niv_5										AS Cod_Niv_5,	")
			loComandoSeleccionar.AppendLine("		(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables									")  
			loComandoSeleccionar.AppendLine("			WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_5)AS Nom_Niv_5,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Niv_6										AS Cod_Niv_6,	") 
			loComandoSeleccionar.AppendLine("		(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables									") 
			loComandoSeleccionar.AppendLine("			WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_6)AS Nom_Niv_6,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Documento 										AS Documento,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Renglon 										AS Renglon,		") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Referencia										AS Referencia,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Comentario										AS Comentario,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Cue 										AS Cod_Cue, 	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Nom_Cue 										AS Nom_Cue, 	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Cen 										AS Cod_Cen, 	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Gas 										AS Cod_Gas, 	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Aux 										AS Cod_Aux, 	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Nom_Aux 										AS Nom_Aux, 	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Fec_Ori 										AS Fec_Ini, 	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Saldo_Inicial									AS Original,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Saldo_Inicial													") 
			loComandoSeleccionar.AppendLine("			+ ISNULL(SUM(Iniciales.Debe-Iniciales.Haber), 0)			AS Saldo_Inicial,") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Debe											AS Debe,		")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Haber											AS Haber,		")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Saldo_Inicial													")
			loComandoSeleccionar.AppendLine("			+ ISNULL(SUM(Iniciales.Debe-Iniciales.Haber), 0)							") 
			loComandoSeleccionar.AppendLine("			+ #tmpMovimientos.Debe														") 
			loComandoSeleccionar.AppendLine("			- #tmpMovimientos.Haber										AS Saldo_Actual	")
			loComandoSeleccionar.AppendLine("		")
			loComandoSeleccionar.AppendLine("		")
			loComandoSeleccionar.AppendLine("FROM	#tmpMovimientos") 
			loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpMovimientos AS Iniciales") 
			loComandoSeleccionar.AppendLine("		ON	#tmpMovimientos.Cod_Cue = Iniciales.Cod_Cue")
			loComandoSeleccionar.AppendLine("		AND	Iniciales.Posicion < #tmpMovimientos.Posicion ")
			loComandoSeleccionar.AppendLine("GROUP BY	#tmpMovimientos.Posicion,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Documento,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Renglon,")	
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Referencia,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Comentario,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Cue,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Nom_Cue,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Cen,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Gas,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Aux,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Nom_Aux,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Fec_Ori,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Saldo_Inicial,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Debe,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Haber,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_1,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_2,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_3,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_4,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_5,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_6")
			loComandoSeleccionar.AppendLine("ORDER BY	#tmpMovimientos.Cod_Cue, ")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Posicion") 
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes", 450)

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

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMAnalitico_Ampliado_FechaOrigen", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrMAnalitico_Ampliado_FechaOrigen.ReportSource = loObjetoReporte

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
' RJG: 23/06/12: Codigo inicial, a partir de Mayor Analítico Ampliado (se sustituyó la		'
'				 versión original del reporte por otra más rápida.							'
'-------------------------------------------------------------------------------------------' 
' RJG: 29/08/13: Corrección en el saldo final (revisión) de las cuentas de movimiento: debía'
'                coincidir con el último saldo de la cuenta.                        		'
'-------------------------------------------------------------------------------------------' 
