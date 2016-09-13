'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEGanancias_Perdidas_xCentroCostos"
'-------------------------------------------------------------------------------------------'
Partial Class rEGanancias_Perdidas_xCentroCostos
	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
	
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try


			Dim lcFechaDesde 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcFechaHasta 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcCuentaContableDesde 	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcCuentaContableHasta 	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
			Dim lcCentroCostoDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcCentroCostoHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcCuentaGastoDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcCuentaGastoHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			Dim lcAuxiliarDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcAuxiliarHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcMonedaDesde 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcMonedaHasta 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcTipoComprobante		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			'Dim lcSoloConMovimiento 	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
			'Dim lcNivelMostrar			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
			
 			Dim llSoloMovimientos As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(7)).Trim().ToUpper().Equals("SI")
		
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim lcActivo  As String = Strings.Trim(goOpciones.mObtener("CODCUEACT", ""))
			Dim lcPasivo  As String = Strings.Trim(goOpciones.mObtener("CODCUEPAS", ""))
			Dim lcCapital As String = Strings.Trim(goOpciones.mObtener("CODCUECAP", ""))
			
			If (lcActivo="") OrElse (lcPasivo = "") OrElse (lcCapital = "") Then 
				
				Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Operación no Completada", _
					"Debe establecer la opciones 'CODCUEACT', 'CODCUEPAS' y 'CODCUECAP' para indicar cuales son las cuentas" & _
					" raiz del Activo, Pasivo y Capital respectivamente.", _
					 vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Advertencia, _
					 "auto", _
					 "auto")
				Return 
				
			End If
			
			lcActivo	= goServicios.mObtenerCampoFormatoSQL(lcActivo  & "%")
			lcPasivo	= goServicios.mObtenerCampoFormatoSQL(lcPasivo  & "%")
			lcCapital	= goServicios.mObtenerCampoFormatoSQL(lcCapital & "%")



			Dim loConsulta As New StringBuilder()
			Dim lnTamaño1 As Integer = CInt(goOpciones.mObtener("ANCNIV1", "")) 
			Dim lnTamaño2 As Integer = CInt(goOpciones.mObtener("ANCNIV2", "")) 
			Dim lnTamaño3 As Integer = CInt(goOpciones.mObtener("ANCNIV3", "")) 
			Dim lnTamaño4 As Integer = CInt(goOpciones.mObtener("ANCNIV4", "")) 
			Dim lnTamaño5 As Integer = CInt(goOpciones.mObtener("ANCNIV5", "")) 
			Dim lnTamaño6 As Integer = CInt(goOpciones.mObtener("ANCNIV6", ""))
			
			Dim lnCantidad As Integer = CInt(goOpciones.mObtener("CANNIVCUE", ""))				'Niveles usados por la empresa
			Dim lnNivelMax As Integer = CInt(cusAplicacion.goReportes.paParametrosIniciales(8))	'Nivel máximo a mostrar	en reporte
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


			loConsulta.AppendLine("") 
			loConsulta.AppendLine("")
			loConsulta.AppendLine("DECLARE @lcActivo VARCHAR(30)")
			loConsulta.AppendLine("DECLARE @lcPasivo VARCHAR(30)")
			loConsulta.AppendLine("DECLARE @lcCapital VARCHAR(30)")
			loConsulta.AppendLine("DECLARE @lnCero DECIMAL(28, 10)")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("SET	@lcActivo = "	& lcActivo)
			loConsulta.AppendLine("SET	@lcPasivo = "	& lcPasivo)
			loConsulta.AppendLine("SET	@lcCapital = "	& lcCapital)
			loConsulta.AppendLine("SET	@lnCero = 0")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("DECLARE @lnTamaño1 INT; SET @lnTamaño1= " & lnTamaño1.ToString() & ";")
			loConsulta.AppendLine("DECLARE @lnTamaño2 INT; SET @lnTamaño2= " & lnTamaño2.ToString() & ";")
			loConsulta.AppendLine("DECLARE @lnTamaño3 INT; SET @lnTamaño3= " & lnTamaño3.ToString() & ";")
			loConsulta.AppendLine("DECLARE @lnTamaño4 INT; SET @lnTamaño4= " & lnTamaño4.ToString() & ";")
			loConsulta.AppendLine("DECLARE @lnTamaño5 INT; SET @lnTamaño5= " & lnTamaño5.ToString() & ";")
			loConsulta.AppendLine("DECLARE @lnTamaño6 INT; SET @lnTamaño6= " & lnTamaño6.ToString() & ";")
			loConsulta.AppendLine("DECLARE @llUsarTasa BIT; SET @llUsarTasa = 0;")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("DECLARE @lnNivelMax AS INT;")
			loConsulta.AppendLine("SET @lnNivelMax = " & lnNivelMax.ToString())
			loConsulta.AppendLine("DECLARE @lnLongMaxima AS INT;")
			loConsulta.AppendLine("SET @lnLongMaxima = " & lnLongMax.ToString())
			loConsulta.AppendLine("")
			loConsulta.AppendLine("") 
			loConsulta.AppendLine("") 
			loConsulta.AppendLine("") 
			loConsulta.AppendLine("SELECT	SUBSTRING(CC.Cod_Cue, 1, @lnLongMaxima)	AS Cod_Cue,") 
			loConsulta.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño1 AND @lnNivelMax > 1) THEN @lnTamaño1 ELSE 0 END)	AS Cod_Niv_1,") 
			loConsulta.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño2 AND @lnNivelMax > 2) THEN @lnTamaño2 ELSE 0 END)	AS Cod_Niv_2,") 
			loConsulta.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño3 AND @lnNivelMax > 3) THEN @lnTamaño3 ELSE 0 END)	AS Cod_Niv_3,") 
			loConsulta.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño4 AND @lnNivelMax > 4) THEN @lnTamaño4 ELSE 0 END)	AS Cod_Niv_4,") 
			loConsulta.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño5 AND @lnNivelMax > 5) THEN @lnTamaño5 ELSE 0 END)	AS Cod_Niv_5,") 
			loConsulta.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño6 AND @lnNivelMax > 6) THEN @lnTamaño6 ELSE 0 END)	AS Cod_Niv_6,") 
			loConsulta.AppendLine("		COALESCE(Renglones_Comprobantes.Cod_Cen, '')                                                                            AS Cod_Cen,") 
			loConsulta.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < " & lcFechaDesde & ")") 
			loConsulta.AppendLine("					THEN COALESCE(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN COALESCE(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loConsulta.AppendLine("					ELSE @lnCero") 
			loConsulta.AppendLine("			END") 
			loConsulta.AppendLine("		) AS saldo,") 
			loConsulta.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcFechaDesde & " AND Renglones_Comprobantes.Fec_Ini<=" & lcFechaHasta & ")") 
			loConsulta.AppendLine("					THEN COALESCE(Renglones_Comprobantes.Mon_Deb, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN COALESCE(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loConsulta.AppendLine("					ELSE @lnCero") 
			loConsulta.AppendLine("			END") 
			loConsulta.AppendLine("		) AS Debe,") 
			loConsulta.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcFechaDesde & " AND Renglones_Comprobantes.Fec_Ini<=" & lcFechaHasta & ")") 
			loConsulta.AppendLine("					THEN COALESCE(Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN COALESCE(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loConsulta.AppendLine("					ELSE @lnCero") 
			loConsulta.AppendLine("			END") 
			loConsulta.AppendLine("		) AS Haber,") 
			loConsulta.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcFechaDesde & " AND Renglones_Comprobantes.Fec_Ini<=" & lcFechaHasta & ")") 
			loConsulta.AppendLine("					THEN COALESCE(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN COALESCE(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loConsulta.AppendLine("					ELSE @lnCero") 
			loConsulta.AppendLine("			END") 
			loConsulta.AppendLine("		) AS Monto")   
			loConsulta.AppendLine("INTO	#tmpMovimientos") 
			loConsulta.AppendLine("FROM	Cuentas_Contables AS CC") 
			loConsulta.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ") 
			loConsulta.AppendLine("		INNER JOIN Comprobantes") 
			loConsulta.AppendLine("			ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ") 
			loConsulta.AppendLine("				AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
			loConsulta.AppendLine("				AND Comprobantes.Tipo = " & lcTipoComprobante & " AND Comprobantes.Status <> 'Anulado'") 
			loConsulta.AppendLine("			)")
			loConsulta.AppendLine("		ON CC.Cod_Cue = Renglones_Comprobantes.Cod_Cue ") 
			loConsulta.AppendLine("			AND (Renglones_Comprobantes.fec_ini <= " & lcFechaHasta & ")") 
			loConsulta.AppendLine("WHERE	CC.Movimiento=1") 
			loConsulta.AppendLine("		AND	CC.Cod_Cue						BETWEEN " & lcCuentaContableDesde	& "	AND	" & lcCuentaContableHasta)
			loConsulta.AppendLine("		AND Renglones_Comprobantes.Cod_Cen	BETWEEN " & lcCentroCostoDesde		& "	AND	" & lcCentroCostoHasta)
			loConsulta.AppendLine("		AND Renglones_Comprobantes.Cod_Gas	BETWEEN " & lcCuentaGastoDesde		& "	AND	" & lcCuentaGastoHasta)
			loConsulta.AppendLine("		AND Renglones_Comprobantes.Cod_Aux	BETWEEN " & lcAuxiliarDesde			& "	AND	" & lcAuxiliarHasta)
			loConsulta.AppendLine("		AND Renglones_Comprobantes.Cod_Mon	BETWEEN " & lcMonedaDesde			& "	AND	" & lcMonedaHasta)
			loConsulta.AppendLine("GROUP BY CC.Cod_Cue, Renglones_Comprobantes.Cod_Cen") 
			loConsulta.AppendLine("") 
			
			If (lcSeparador <> "") Then
				Dim lcSeparadorSQL  As String = goServicios.mObtenerCampoFormatoSQL("%" & lcSeparador)
				loConsulta.AppendLine("") 
				loConsulta.AppendLine("UPDATE	#tmpMovimientos") 
				loConsulta.AppendLine("SET		Cod_Cue = SUBSTRING(Cod_Cue, 1, LEN(Cod_Cue) - 1)") 
				loConsulta.AppendLine("WHERE	COALESCE(Cod_Cue, '') LIKE " & lcSeparadorSQL ) 
				loConsulta.AppendLine("") 
				loConsulta.AppendLine("UPDATE	#tmpMovimientos") 
				loConsulta.AppendLine("SET		Cod_Niv_1 = SUBSTRING(Cod_Niv_1, 1, LEN(Cod_Niv_1) - 1)") 
				loConsulta.AppendLine("WHERE	COALESCE(Cod_Niv_1, '') LIKE " & lcSeparadorSQL )
				loConsulta.AppendLine("") 
				loConsulta.AppendLine("UPDATE	#tmpMovimientos									") 
				loConsulta.AppendLine("SET		Cod_Niv_2 = SUBSTRING(Cod_Niv_2, 1, LEN(Cod_Niv_2) - 1)") 
				loConsulta.AppendLine("WHERE	COALESCE(Cod_Niv_2, '') LIKE " & lcSeparadorSQL ) 
				loConsulta.AppendLine("") 
				loConsulta.AppendLine("UPDATE	#tmpMovimientos									") 
				loConsulta.AppendLine("SET		Cod_Niv_3 = SUBSTRING(Cod_Niv_3, 1, LEN(Cod_Niv_3) - 1)") 
				loConsulta.AppendLine("WHERE	COALESCE(Cod_Niv_3, '') LIKE " & lcSeparadorSQL ) 
				loConsulta.AppendLine("") 
				loConsulta.AppendLine("UPDATE	#tmpMovimientos									") 
				loConsulta.AppendLine("SET		Cod_Niv_4 = SUBSTRING(Cod_Niv_4, 1, LEN(Cod_Niv_4) - 1)") 
				loConsulta.AppendLine("WHERE	COALESCE(Cod_Niv_4, '') LIKE " & lcSeparadorSQL ) 
				loConsulta.AppendLine("") 
				loConsulta.AppendLine("UPDATE	#tmpMovimientos									") 
				loConsulta.AppendLine("SET		Cod_Niv_5 = SUBSTRING(Cod_Niv_5, 1, LEN(Cod_Niv_5) - 1)") 
				loConsulta.AppendLine("WHERE	COALESCE(Cod_Niv_5, '') LIKE " & lcSeparadorSQL ) 
				loConsulta.AppendLine("") 
				loConsulta.AppendLine("UPDATE	#tmpMovimientos									") 
				loConsulta.AppendLine("SET		Cod_Niv_6 = SUBSTRING(Cod_Niv_6, 1, LEN(Cod_Niv_6) - 1)") 
				loConsulta.AppendLine("WHERE	COALESCE(Cod_Niv_6, '') LIKE " & lcSeparadorSQL ) 
				loConsulta.AppendLine("") 
			End If
			loConsulta.AppendLine("") 
			loConsulta.AppendLine("") 
			loConsulta.AppendLine("SELECT	Cuentas_Contables.Cod_Cue							AS Cod_Cue,")
			loConsulta.AppendLine("			Cuentas_Contables.Nom_Cue							AS Nom_Cue,")
			loConsulta.AppendLine("			Cuentas_Contables.Movimiento						AS Movimiento,")
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Cen, ")
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_1, ")
			loConsulta.AppendLine("			(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_1) AS Nom_Niv_1,")
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_2, ")
			loConsulta.AppendLine("			(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_2) AS Nom_Niv_2,")
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_3, ")
			loConsulta.AppendLine("			(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_3) AS Nom_Niv_3,")
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_4, ")
			loConsulta.AppendLine("			(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_4) AS Nom_Niv_4,")
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_5, ")
			loConsulta.AppendLine("			(SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_5) AS Nom_Niv_5,")
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_6, (")
			loConsulta.AppendLine("			SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_6) AS Nom_Niv_6,")
			loConsulta.AppendLine("			SUM(#tmpMovimientos.Saldo)							AS Saldo_Inicial,") 
			loConsulta.AppendLine("			SUM(#tmpMovimientos.Debe)							AS Debe,") 
			loConsulta.AppendLine("			SUM(#tmpMovimientos.Haber)							AS Haber,") 
			loConsulta.AppendLine("			SUM(#tmpMovimientos.Monto + #tmpMovimientos.Saldo)	AS Saldo_Actual") 
			loConsulta.AppendLine("FROM		#tmpMovimientos") 
			loConsulta.AppendLine("	JOIN	Cuentas_Contables") 
			loConsulta.AppendLine("		ON	Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Cue") 
			loConsulta.AppendLine("WHERE		(		NOT Cuentas_Contables.Cod_Cue LIKE @lcActivo ")
			loConsulta.AppendLine("				AND	NOT Cuentas_Contables.Cod_Cue LIKE @lcPasivo ")
			loConsulta.AppendLine("				AND	NOT Cuentas_Contables.Cod_Cue LIKE @lcCapital ")
			loConsulta.AppendLine("			)")
			loConsulta.AppendLine("GROUP BY	Cuentas_Contables.Cod_Cue, ") 
			loConsulta.AppendLine("			Cuentas_Contables.Nom_Cue, ") 
			loConsulta.AppendLine("			Cuentas_Contables.Movimiento,")
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Cen, ") 
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_1, ") 
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_2, ")
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_3,") 
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_4, ")
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_5, ")
			loConsulta.AppendLine("			#tmpMovimientos.Cod_Niv_6")
			If llSoloMovimientos Then
				loConsulta.AppendLine("HAVING	    ABS(SUM(#tmpMovimientos.Saldo)) + ABS(SUM(#tmpMovimientos.Debe)) + ABS(SUM(#tmpMovimientos.Haber)) > 0") 
			End If
			loConsulta.AppendLine("ORDER BY	" & lcOrdenamiento) 
			loConsulta.AppendLine("") 
			loConsulta.AppendLine("DROP TABLE #tmpMovimientos") 
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")
			loConsulta.AppendLine("")

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEGanancias_Perdidas_xCentroCostos", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrEGanancias_Perdidas_xCentroCostos.ReportSource = loObjetoReporte

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
' RJG: 26/10/11: Codigo inicial, a partir de rEGanancias_Perdidas.'
'-------------------------------------------------------------------------------------------' 
