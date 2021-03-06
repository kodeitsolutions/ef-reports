﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEGanancias_Perdidas_Mensual"
'-------------------------------------------------------------------------------------------'
Partial Class rEGanancias_Perdidas_Mensual

	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
	
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try


			Dim lnAno        			As Integer = cusAplicacion.goReportes.paParametrosIniciales(0)
			'Dim lcFechaHasta 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			'Dim lcPrimerSemestre 	    As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcCuentaContableDesde 	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcCuentaContableHasta 	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcCentroCostoDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcCentroCostoHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			Dim lcCuentaGastoDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcCuentaGastoHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcAuxiliarDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcAuxiliarHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			Dim lcMonedaDesde 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcMonedaHasta 			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
			Dim lcTipoComprobante		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
			'Dim lcSoloConMovimiento 	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
			'Dim lcNivelMostrar			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
			
 			Dim llPrimerSemestre As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(1)).Trim().ToUpper().Equals("SI")
 			Dim llSoloMovimientos As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(8)).Trim().ToUpper().Equals("SI")
		
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
			Dim lnNivelMax As Integer = CInt(cusAplicacion.goReportes.paParametrosIniciales(9))	'Nivel máximo a mostrar	en reporte
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
            
            Dim ldInicio As Date 
            if (lnAno < 1900) Then 
                lnAno = Date.Today.Year
            End If
            If llPrimerSemestre Then 
                ldInicio = New Date(lnAno, 1, 1)
            Else
                ldInicio = New Date(lnAno, 7, 1)
            End If
            Dim lcCorte00 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio)
            Dim lcCorte01 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(1))
            Dim lcCorte02 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(2))
            Dim lcCorte03 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(3))
            Dim lcCorte04 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(4))
            Dim lcCorte05 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(5))
            Dim lcCorte06 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(6))


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
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < " & lcCorte00 & ")") 
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS saldo,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte00 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte01 & ")") 
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto01,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte01 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte02 & ")") 
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto02,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte02 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte03 & ")") 
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto03,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte03 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte04 & ")") 
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto04,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte04 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte05 & ")") 
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto05,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte05 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte06 & ")") 
			loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto06")   
			loComandoSeleccionar.AppendLine("INTO	#tmpMovimientos") 
			loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables AS CC") 
			loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ") 
			loComandoSeleccionar.AppendLine("		INNER JOIN Comprobantes") 
			loComandoSeleccionar.AppendLine("			ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ") 
			loComandoSeleccionar.AppendLine("				AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
			loComandoSeleccionar.AppendLine("				AND Comprobantes.Tipo = " & lcTipoComprobante & " AND Comprobantes.Status <> 'Anulado'") 
			loComandoSeleccionar.AppendLine("			)")
			loComandoSeleccionar.AppendLine("		ON CC.Cod_Cue = Renglones_Comprobantes.Cod_Cue ") 
			loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.fec_ini < " & lcCorte06 & ")") 
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
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto01)	                    AS Monto_01,") 
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto02)	                    AS Monto_02,") 
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto03)	                    AS Monto_03,")  
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto04)	                    AS Monto_04,")
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto05)	                    AS Monto_05,") 
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto06)	                    AS Monto_06,") 
			loComandoSeleccionar.AppendLine("			(SUM(#tmpMovimientos.Monto01)+SUM(#tmpMovimientos.Monto02)+SUM(#tmpMovimientos.Monto03)+")     
			loComandoSeleccionar.AppendLine("			 SUM(#tmpMovimientos.Monto04)+SUM(#tmpMovimientos.Monto05)+SUM(#tmpMovimientos.Monto06)) AS Acumulado,") 
			loComandoSeleccionar.AppendLine("			" & IIf(llPrimerSemestre, "1", "2") & "	                    AS Semestre") 
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
				loComandoSeleccionar.AppendLine("HAVING	    ABS(SUM(#tmpMovimientos.Monto01))+ABS(SUM(#tmpMovimientos.Monto02))+ABS(SUM(#tmpMovimientos.Monto03))") 
				loComandoSeleccionar.AppendLine("          + ABS(SUM(#tmpMovimientos.Monto04))+ABS(SUM(#tmpMovimientos.Monto05))+ABS(SUM(#tmpMovimientos.Monto06)) > 0") 
			End If
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento) 
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos") 
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			
			Dim loServicios As New cusDatos.goDatos
            
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEGanancias_Perdidas_Mensual", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrEGanancias_Perdidas_Mensual.ReportSource = loObjetoReporte

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
' RJG: 16/06/15: Codigo inicial, a partir de rEGanancias_Perdidas.						    '
'-------------------------------------------------------------------------------------------' 
