'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEGanancias_Perdidas"
'-------------------------------------------------------------------------------------------'
Partial Class rEGanancias_Perdidas

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



			Dim loComandoSeleccionar As New StringBuilder()
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


			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lcActivo VARCHAR(30)")
			loComandoSeleccionar.AppendLine("DECLARE @lcPasivo VARCHAR(30)")
			loComandoSeleccionar.AppendLine("DECLARE @lcCapital VARCHAR(30)")
			loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10)")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SET	@lcActivo = "	& lcActivo)
			loComandoSeleccionar.AppendLine("SET	@lcPasivo = "	& lcPasivo)
			loComandoSeleccionar.AppendLine("SET	@lcCapital = "	& lcCapital)
			loComandoSeleccionar.AppendLine("SET	@lnCero = 0")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño1 INT; SET @lnTamaño1= " & lnTamaño1.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño2 INT; SET @lnTamaño2= " & lnTamaño2.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño3 INT; SET @lnTamaño3= " & lnTamaño3.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño4 INT; SET @lnTamaño4= " & lnTamaño4.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño5 INT; SET @lnTamaño5= " & lnTamaño5.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño6 INT; SET @lnTamaño6= " & lnTamaño6.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @llUsarTasa BIT; SET @llUsarTasa = 0;")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lnNivelMax AS INT;")
			loComandoSeleccionar.AppendLine("SET @lnNivelMax = " & lnNivelMax.ToString())
			loComandoSeleccionar.AppendLine("DECLARE @lnLongMaxima AS INT;")
			loComandoSeleccionar.AppendLine("SET @lnLongMaxima = " & lnLongMax.ToString())
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("SELECT	SUBSTRING(CC.Cod_Cue, 1, @lnLongMaxima)	AS Cod_Cue,") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño1 AND @lnNivelMax > 1) THEN @lnTamaño1 ELSE 0 END)	AS Cod_Niv_1,") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño2 AND @lnNivelMax > 2) THEN @lnTamaño2 ELSE 0 END)	AS Cod_Niv_2,") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño3 AND @lnNivelMax > 3) THEN @lnTamaño3 ELSE 0 END)	AS Cod_Niv_3,") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño4 AND @lnNivelMax > 4) THEN @lnTamaño4 ELSE 0 END)	AS Cod_Niv_4,") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño5 AND @lnNivelMax > 5) THEN @lnTamaño5 ELSE 0 END)	AS Cod_Niv_5,") 
			loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño6 AND @lnNivelMax > 6) THEN @lnTamaño6 ELSE 0 END)	AS Cod_Niv_6,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < " & lcFechaDesde & ")") 
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS saldo,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcFechaDesde & " AND Renglones_Comprobantes.Fec_Ini<=" & lcFechaHasta & ")") 
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Debe,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcFechaDesde & " AND Renglones_Comprobantes.Fec_Ini<=" & lcFechaHasta & ")") 
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Haber,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcFechaDesde & " AND Renglones_Comprobantes.Fec_Ini<=" & lcFechaHasta & ")") 
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto")   
			loComandoSeleccionar.AppendLine("INTO	#tmpMovimientos") 
			loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables AS CC") 
			loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ") 
			loComandoSeleccionar.AppendLine("		INNER JOIN Comprobantes") 
			loComandoSeleccionar.AppendLine("			ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ") 
			loComandoSeleccionar.AppendLine("				AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
			loComandoSeleccionar.AppendLine("				AND Comprobantes.Tipo = " & lcTipoComprobante & " AND Comprobantes.Status <> 'Anulado'") 
			loComandoSeleccionar.AppendLine("			)")
			loComandoSeleccionar.AppendLine("		ON CC.Cod_Cue = Renglones_Comprobantes.Cod_Cue ") 
			loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.fec_ini <= " & lcFechaHasta & ")") 
			loComandoSeleccionar.AppendLine("WHERE	CC.Movimiento=1") 
			loComandoSeleccionar.AppendLine("		AND	CC.Cod_Cue						BETWEEN " & lcCuentaContableDesde	& "	AND	" & lcCuentaContableHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Cen	BETWEEN " & lcCentroCostoDesde		& "	AND	" & lcCentroCostoHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Gas	BETWEEN " & lcCuentaGastoDesde		& "	AND	" & lcCuentaGastoHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Aux	BETWEEN " & lcAuxiliarDesde			& "	AND	" & lcAuxiliarHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Mon	BETWEEN " & lcMonedaDesde			& "	AND	" & lcMonedaHasta)
			loComandoSeleccionar.AppendLine("GROUP BY CC.Cod_Cue") 
			loComandoSeleccionar.AppendLine("") 
			
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
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("SELECT		Cuentas_Contables.Cod_Cue							AS Cod_Cue,")
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Nom_Cue							AS Nom_Cue,")
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Movimiento						AS Movimiento,")
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
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_6, (")
			loComandoSeleccionar.AppendLine("			SELECT TOP 1 Nom_Cue FROM Cuentas_Contables WHERE Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Niv_6) AS Nom_Niv_6,")
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Saldo)							AS Saldo_Inicial,") 
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Debe)							AS Debe,") 
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Haber)							AS Haber,") 
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto + #tmpMovimientos.Saldo)	AS Saldo_Actual") 
			loComandoSeleccionar.AppendLine("FROM		#tmpMovimientos") 
			loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Contables") 
			loComandoSeleccionar.AppendLine("		ON	Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Cue") 
			loComandoSeleccionar.AppendLine("WHERE		(		NOT Cuentas_Contables.Cod_Cue LIKE @lcActivo ")
			loComandoSeleccionar.AppendLine("				AND	NOT Cuentas_Contables.Cod_Cue LIKE @lcPasivo ")
			loComandoSeleccionar.AppendLine("				AND	NOT Cuentas_Contables.Cod_Cue LIKE @lcCapital ")
			loComandoSeleccionar.AppendLine("			)")
			loComandoSeleccionar.AppendLine("GROUP BY	Cuentas_Contables.Cod_Cue, ") 
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Nom_Cue, ") 
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Movimiento,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_1, ") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_2, ")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_3,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_4, ")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_5, ")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_6")
			If llSoloMovimientos Then
				loComandoSeleccionar.AppendLine("HAVING	ABS(SUM(#tmpMovimientos.Saldo)) + ABS(SUM(#tmpMovimientos.Debe)) + ABS(SUM(#tmpMovimientos.Haber)) > 0") 
			End If
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento) 
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos") 
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
																													
			'Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			'Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			'Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			'Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
			'Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			'Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			'Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			'Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			'Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			'Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			'Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			'Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			'Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			'Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
			'Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
			
			'Dim llSoloMovimientos As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(7)).Trim().ToUpper().Equals("SI")
		
			'Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			'Dim loComandoSeleccionar As New StringBuilder()

			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Desde DATETIME	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Hasta DATETIME	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Desde VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Hasta VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Desde VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Hasta VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Desde VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Hasta VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Desde VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Hasta VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro6Hasta VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro6Desde VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro7Desde VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro8Desde VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("DECLARE @lcParametro9Desde VARCHAR(100)	")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("SET		@lcParametro1Desde = " & lcParametro0Desde)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro1Hasta = " & lcParametro0Hasta)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro2Desde = " & lcParametro1Desde)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro2Hasta = " & lcParametro1Hasta)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro3Desde = " & lcParametro2Desde)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro3Hasta = " & lcParametro2Hasta)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro4Desde = " & lcParametro3Desde)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro4Hasta = " & lcParametro3Hasta)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro5Desde = " & lcParametro4Desde)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro5Hasta = " & lcParametro4Hasta)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro6Desde = " & lcParametro5Desde)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro6Hasta = " & lcParametro5Hasta)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro7Desde = " & lcParametro6Desde)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro8Desde = " & lcParametro7Desde)
			'loComandoSeleccionar.AppendLine("SET		@lcParametro9Desde = " & lcParametro8Desde)
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10)")
			'loComandoSeleccionar.AppendLine("DECLARE @llFalso BIT")
			'loComandoSeleccionar.AppendLine("DECLARE @llVerdadero BIT")		 
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("SET	@lnCero = 0")
			'loComandoSeleccionar.AppendLine("SET	@llFalso = 0")
			'loComandoSeleccionar.AppendLine("SET	@llVerdadero = 1")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("-- Prepara un listado de las cuentas contables a incluir  *")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("SELECT		LEN(RTRIM(Cod_Cue))				AS Nivel, ")
			'loComandoSeleccionar.AppendLine("			CAST(Cod_Cue AS VARCHAR(100))	AS Cod_cue, ")
			'loComandoSeleccionar.AppendLine("			CAST(Nom_Cue AS VARCHAR(100))	AS Nom_Cue, ")
			'loComandoSeleccionar.AppendLine("			Movimiento						AS Movimiento,")
			'loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_1,	")
			'loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_2,	")
			'loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_3,	")
			'loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_4,	")
			'loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_5,	")
			'loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_6,	")
			'loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_7,	")
			'loComandoSeleccionar.AppendLine("			Mon_Ini							AS Saldo_Inicial")
			'loComandoSeleccionar.AppendLine("INTO		#tmpCuentas")
			'loComandoSeleccionar.AppendLine("FROM		Cuentas_Contables")
			'loComandoSeleccionar.AppendLine("WHERE		Cod_Cue BETWEEN @lcParametro2Desde AND @lcParametro2Hasta")
			'loComandoSeleccionar.AppendLine("	AND		Cod_Cue >= '4'")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("-- Agrega los saldos iniciales.							  *")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("UPDATE		#tmpCuentas")
			'loComandoSeleccionar.AppendLine("SET			#tmpCuentas.Saldo_Inicial = #tmpCuentas.Saldo_Inicial ")
			'loComandoSeleccionar.AppendLine("										+ Saldos.TotalDebe ")
			'loComandoSeleccionar.AppendLine("										- Saldos.Total_Haber")
			'loComandoSeleccionar.AppendLine("FROM	(	SELECT	RC.Cod_Cue			AS Cod_cue, ")
			'loComandoSeleccionar.AppendLine("					SUM(RC.Mon_Deb)		AS TotalDebe,")
			'loComandoSeleccionar.AppendLine("					SUM(RC.Mon_Hab)		AS Total_Haber")
			'loComandoSeleccionar.AppendLine("			FROM	Renglones_Comprobantes AS RC")
			'loComandoSeleccionar.AppendLine("				JOIN Comprobantes On Comprobantes.Documento = RC.Documento")
			'loComandoSeleccionar.AppendLine("					AND Comprobantes.Adicional = RC.Adicional ")
			'loComandoSeleccionar.AppendLine("					AND Comprobantes.Tipo = @lcParametro7Desde")
			'loComandoSeleccionar.AppendLine("					AND Comprobantes.Status <> 'Anulado'")
			'loComandoSeleccionar.AppendLine("			WHERE	RC.Fec_Ini < @lcParametro1Desde")
			'loComandoSeleccionar.AppendLine("				AND RC.Cod_Cen	BETWEEN @lcParametro3Desde AND @lcParametro3Hasta")
			'loComandoSeleccionar.AppendLine("				AND RC.Cod_Gas	BETWEEN @lcParametro4Desde AND @lcParametro4Hasta")
			'loComandoSeleccionar.AppendLine("				AND RC.Cod_Aux	BETWEEN @lcParametro5Desde AND @lcParametro5Hasta")
			'loComandoSeleccionar.AppendLine("				AND RC.Cod_Mon	BETWEEN @lcParametro6Desde AND @lcParametro6Hasta")
			'loComandoSeleccionar.AppendLine("			")
			'loComandoSeleccionar.AppendLine("			")
			'loComandoSeleccionar.AppendLine("			GROUP BY RC.Cod_Cue")
			'loComandoSeleccionar.AppendLine("		) AS Saldos")
			'loComandoSeleccionar.AppendLine("WHERE	#tmpCuentas.Cod_Cue = Saldos.Cod_Cue")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("-- Calcula el 'rango' de cada nivel.                      *")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("UPDATE		#tmpCuentas")
			'loComandoSeleccionar.AppendLine("SET		Nivel_1 = (CASE WHEN N.Margen > 1 THEN 1 ELSE 0 END),")
			'loComandoSeleccionar.AppendLine("			Nivel_2 = (CASE WHEN N.Margen > 2 THEN 1 ELSE 0 END),")
			'loComandoSeleccionar.AppendLine("			Nivel_3 = (CASE WHEN N.Margen > 3 THEN 1 ELSE 0 END),")
			'loComandoSeleccionar.AppendLine("			Nivel_4 = (CASE WHEN N.Margen > 4 THEN 1 ELSE 0 END),")
			'loComandoSeleccionar.AppendLine("			Nivel_5 = (CASE WHEN N.Margen > 5 THEN 1 ELSE 0 END),")
			'loComandoSeleccionar.AppendLine("			Nivel_6 = (CASE WHEN N.Margen > 6 THEN 1 ELSE 0 END),")
			'loComandoSeleccionar.AppendLine("			Nivel_7 = (CASE WHEN N.Margen > 7 THEN 1 ELSE 0 END),")
			'loComandoSeleccionar.AppendLine("			Nivel = N.Margen			")
			'loComandoSeleccionar.AppendLine("FROM		(	SELECT	ROW_NUMBER() OVER(ORDER BY #tmpCuentas.Nivel ASC) AS Margen,")
			'loComandoSeleccionar.AppendLine("						#tmpCuentas.Nivel AS Original")
			'loComandoSeleccionar.AppendLine("				FROM	#tmpCuentas")
			'loComandoSeleccionar.AppendLine("				GROUP BY #tmpCuentas.Nivel")
			'loComandoSeleccionar.AppendLine("			) AS N")
			'loComandoSeleccionar.AppendLine("WHERE		#tmpCuentas.Nivel = N.Original")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("-- Union principal.                                       *")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("SELECT		#tmpCuentas.Nivel								AS Nivel,")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Cod_Cue								AS Cod_Cue, ")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Nom_Cue								AS Nom_Cue, ")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Movimiento							AS Movimiento,")
			'loComandoSeleccionar.AppendLine("			SUM(ISNULL(tmpRenglones.Mon_Deb, @lnCero))		AS Mon_Deb, ")
			'loComandoSeleccionar.AppendLine("			SUM(ISNULL(tmpRenglones.Mon_Hab, @lnCero))		AS Mon_Hab, ")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_1								AS Nivel_1,")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_2								AS Nivel_2,")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_3								AS Nivel_3,")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_4								AS Nivel_4,")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_5								AS Nivel_5,")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_6								AS Nivel_6,")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_7								AS Nivel_7,")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Saldo_Inicial					AS Saldo_Inicial,")
			'loComandoSeleccionar.AppendLine("			(	SUM(ISNULL(tmpRenglones.Mon_Deb, @lnCero)) -				")
			'loComandoSeleccionar.AppendLine("				SUM(ISNULL(tmpRenglones.Mon_Hab, @lnCero)) +				")
			'loComandoSeleccionar.AppendLine("				#tmpCuentas.Saldo_Inicial)				AS Saldo_Actual")
			'loComandoSeleccionar.AppendLine("INTO		#tmpTodos")
			'loComandoSeleccionar.AppendLine("FROM		#tmpCuentas")
			'loComandoSeleccionar.AppendLine("	LEFT JOIN (")
			'loComandoSeleccionar.AppendLine("				SELECT	Renglones_Comprobantes.Mon_Deb,")
			'loComandoSeleccionar.AppendLine("						Renglones_Comprobantes.Mon_Hab,")
			'loComandoSeleccionar.AppendLine("						Renglones_Comprobantes.Cod_cue")
			'loComandoSeleccionar.AppendLine("				FROM	Renglones_Comprobantes")
			'loComandoSeleccionar.AppendLine("					JOIN	Comprobantes")
			'loComandoSeleccionar.AppendLine("						ON	Comprobantes.Documento = Renglones_Comprobantes.Documento")
			'loComandoSeleccionar.AppendLine("						AND Comprobantes.Adicional = Renglones_Comprobantes.Adicional ")
			'loComandoSeleccionar.AppendLine("						AND Comprobantes.Tipo = @lcParametro7Desde")
			'loComandoSeleccionar.AppendLine("					AND Comprobantes.Status <> 'Anulado'")
			'loComandoSeleccionar.AppendLine("				WHERE		Renglones_Comprobantes.Fec_Ini	BETWEEN @lcParametro1Desde AND @lcParametro1Hasta")
			'loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Cen	BETWEEN @lcParametro3Desde AND @lcParametro3Hasta")
			'loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Gas	BETWEEN @lcParametro4Desde AND @lcParametro4Hasta")
			'loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Aux	BETWEEN @lcParametro5Desde AND @lcParametro5Hasta")
			'loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Mon	BETWEEN @lcParametro6Desde AND @lcParametro6Hasta")
			'loComandoSeleccionar.AppendLine("			) AS tmpRenglones ")
			'loComandoSeleccionar.AppendLine("		ON tmpRenglones.Cod_Cue = #tmpCuentas.Cod_Cue")
			'loComandoSeleccionar.AppendLine("GROUP BY	#tmpCuentas.Nivel, #tmpCuentas.Cod_cue, #tmpCuentas.Nom_Cue, ")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Movimiento, #tmpCuentas.Saldo_Inicial, ")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_1, #tmpCuentas.Nivel_2, #tmpCuentas.Nivel_3, ")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_4, #tmpCuentas.Nivel_5, #tmpCuentas.Nivel_6, ")
			'loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_7")
			'loComandoSeleccionar.AppendLine("ORDER BY	#tmpCuentas.Cod_cue")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("DROP TABLE #tmpCuentas")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("-- Actualiza los totales hasta el nivel indicado.         *")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("UPDATE	#tmpTodos")
			'loComandoSeleccionar.AppendLine("SET	Mon_Deb = T.Total_Debe,")
			'loComandoSeleccionar.AppendLine("		Mon_Hab = T.Total_Haber,")
			'loComandoSeleccionar.AppendLine("		Saldo_Inicial = T.Total_Inicial,")
			'loComandoSeleccionar.AppendLine("		Saldo_Actual = T.Total_Actual")
			'loComandoSeleccionar.AppendLine("FROM (	SELECT	Total.Cod_cue				AS Cuenta_Padre,")
			'loComandoSeleccionar.AppendLine("				SUM(Interna.Mon_Deb)		AS Total_Debe,")
			'loComandoSeleccionar.AppendLine("				SUM(Interna.Mon_Hab)		AS Total_Haber,")
			'loComandoSeleccionar.AppendLine("				SUM(Interna.Saldo_Inicial)	AS Total_Inicial,")
			'loComandoSeleccionar.AppendLine("				SUM(Interna.Saldo_Actual)	AS Total_Actual")
			'loComandoSeleccionar.AppendLine("		FROM	#tmpTodos AS Total")
			'loComandoSeleccionar.AppendLine("			JOIN #tmpTodos AS Interna")
			'loComandoSeleccionar.AppendLine("				ON SUBSTRING(Interna.Cod_Cue, 1, LEN(Total.Cod_Cue)) = Total.Cod_cue")
			'loComandoSeleccionar.AppendLine("				AND Interna.Nivel >= @lcParametro9Desde")
			'loComandoSeleccionar.AppendLine("		GROUP BY Total.Cod_cue")
			'loComandoSeleccionar.AppendLine("	) AS T")
			'loComandoSeleccionar.AppendLine("WHERE #tmpTodos.Nivel = @lcParametro9Desde")
			'loComandoSeleccionar.AppendLine("	AND T.Cuenta_Padre = #tmpTodos.Cod_Cue")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("-- Actualiza el máximo nivel (lo requiere el RPT)         *")
			'loComandoSeleccionar.AppendLine("--*********************************************************")
			'loComandoSeleccionar.AppendLine("DELETE	FROM #tmpTodos")
			'loComandoSeleccionar.AppendLine("WHERE	#tmpTodos.Nivel > @lcParametro9Desde")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("SET @lcParametro9Desde = (SELECT MAX(Nivel) FROM #tmpTodos)")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("")
			'If llSoloMovimientos Then 
			'	loComandoSeleccionar.AppendLine("--*********************************************************")
			'	loComandoSeleccionar.AppendLine("-- Elimina los niveles máximos y sin movimientos          *")
			'	loComandoSeleccionar.AppendLine("--*********************************************************")
			'	loComandoSeleccionar.AppendLine("DELETE		FROM #tmpTodos")
			'	loComandoSeleccionar.AppendLine("WHERE		#tmpTodos.Nivel = @lcParametro9Desde")
			'	loComandoSeleccionar.AppendLine("	AND		(ABS(Saldo_Actual)) = @lnCero")
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("--*********************************************************")
			'	loComandoSeleccionar.AppendLine("-- Elimina los niveles superiores y sin movimientos       *")
			'	loComandoSeleccionar.AppendLine("--*********************************************************")
			'	loComandoSeleccionar.AppendLine("SELECT	@lnCero As SubTotal, Cod_Cue, Nivel")
			'	loComandoSeleccionar.AppendLine("INTO	#tmpTotales")
			'	loComandoSeleccionar.AppendLine("FROM	#tmpTodos")
			'	loComandoSeleccionar.AppendLine("WHERE	#tmpTodos.Nivel < @lcParametro9Desde")
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("DECLARE curNiveles CURSOR FOR")
			'	loComandoSeleccionar.AppendLine("	SELECT DISTINCT Nivel FROM #tmpTotales ORDER BY Nivel DESC")
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("OPEN curNiveles ")
			'	loComandoSeleccionar.AppendLine("DECLARE @lnNivelActual INT")
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("FETCH NEXT FROM curNiveles INTO @lnNivelActual")
			'	loComandoSeleccionar.AppendLine("WHILE @@FETCH_STATUS = 0")
			'	loComandoSeleccionar.AppendLine("BEGIN")
			'	loComandoSeleccionar.AppendLine("		UPDATE		#tmpTotales")		
			'	loComandoSeleccionar.AppendLine("		SET			#tmpTotales.SubTotal = SubTotales.SubTotal")		
			'	loComandoSeleccionar.AppendLine("		FROM		(	SELECT		SUM(A.Saldo_Actual) AS SubTotal,")		
			'	loComandoSeleccionar.AppendLine("									B.Cod_Cue")		
			'	loComandoSeleccionar.AppendLine("						FROM		#tmpTodos AS A")		
			'	loComandoSeleccionar.AppendLine("							JOIN	#tmpTotales AS B ON A.Cod_Cue LIKE (RTRIM(B.Cod_Cue) + '%')")		
			'	loComandoSeleccionar.AppendLine("						WHERE		A.Nivel > B.Nivel")		
			'	loComandoSeleccionar.AppendLine("							AND		B.Nivel = @lnNivelActual")		
			'	loComandoSeleccionar.AppendLine("						GROUP BY	B.Cod_Cue")		
			'	loComandoSeleccionar.AppendLine("					) AS SubTotales")		
			'	loComandoSeleccionar.AppendLine("		WHERE		SubTotales.Cod_Cue = #tmpTotales.Cod_Cue")		
			'	loComandoSeleccionar.AppendLine("			AND		#tmpTotales.Nivel = @lnNivelActual")
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("		FETCH NEXT FROM curNiveles INTO @lnNivelActual")		
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("END")
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("CLOSE curNiveles ")
			'	loComandoSeleccionar.AppendLine("DEALLOCATE curNiveles ")
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("DELETE #tmpTodos")
			'	loComandoSeleccionar.AppendLine("FROM	#tmpTotales As Totales")
			'	loComandoSeleccionar.AppendLine("WHERE	#tmpTodos.Cod_cue = Totales.Cod_Cue")
			'	loComandoSeleccionar.AppendLine("		AND Totales.SubTotal = @lnCero")
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("DROP TABLE #tmpTotales")
			'	loComandoSeleccionar.AppendLine("")
			'	loComandoSeleccionar.AppendLine("")
			'End If
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("SELECT		CAST(@lcParametro9Desde AS INT) AS Nivel_Maximo, ")
			'loComandoSeleccionar.AppendLine("			#tmpTodos.* ")
			'loComandoSeleccionar.AppendLine("FROM		#tmpTodos")
			'loComandoSeleccionar.AppendLine("ORDER BY	Cod_cue")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("DROP TABLE #tmpTodos")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("")
			'loComandoSeleccionar.AppendLine("")
			
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEGanancias_Perdidas", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrEGanancias_Perdidas.ReportSource = loObjetoReporte

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
' RJG: 27/10/11: Corrección en cálculo de saldos inicial y final. Corrección en filtro.	'
'-------------------------------------------------------------------------------------------' 
' RJG: 16/11/11: Agregado filtro de la estructura superior cuando en el detalle no hay		'
'				 movimientos (y el usuario indicó el filtro "Solo Moviminetos = SI".		'
'-------------------------------------------------------------------------------------------' 
' RJG: 06/12/11: Se agregó la igualdad de campo Adicional en las uniones entre Comprobantes	'
'				 y sus renglones.															'
'-------------------------------------------------------------------------------------------' 
' RJG: 09/02/12: Cambiado el filtro de comprobantes "='Pendiente'" por "<>'Anulado'".		'
'-------------------------------------------------------------------------------------------' 
