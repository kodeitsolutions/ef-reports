'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rEGanancias_Perdidas_Mensual"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rEGanancias_Perdidas_Mensual

	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
	
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try


			Dim lnAno        			As Integer = cusAplicacion.goReportes.paParametrosIniciales(0)
            Dim lcCuentaContableDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcCuentaContableHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcAuxiliarDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcAuxiliarHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))

            Dim llSoloMovimientos As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(3)).Trim().ToUpper().Equals("SI")

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
            lcCapital = goServicios.mObtenerCampoFormatoSQL(lcCapital & "%")

            Dim loComandoSeleccionar As New StringBuilder()
            Dim lnNivelMax As Integer = CInt(cusAplicacion.goReportes.paParametrosIniciales(4)) 'Nivel máximo a mostrar	en reporte

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

            Dim ldInicio As Date
            If (lnAno < 1900) Then
                lnAno = Date.Today.Year
            End If
            ldInicio = New Date(lnAno, 1, 1)

            Dim lcCorte00 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio)
            Dim lcCorte01 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(1))
            Dim lcCorte02 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(2))
            Dim lcCorte03 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(3))
            Dim lcCorte04 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(4))
            Dim lcCorte05 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(5))
            Dim lcCorte06 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(6))
            Dim lcCorte07 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(7))
            Dim lcCorte08 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(8))
            Dim lcCorte09 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(9))
            Dim lcCorte10 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(10))
            Dim lcCorte11 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(11))
            Dim lcCorte12 As String = goServicios.mObtenerCampoFormatoSQL(ldInicio.AddMonths(12))

            loComandoSeleccionar.AppendLine("DECLARE @lcCuentaDesde VARCHAR(30) = " & lcCuentaContableDesde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCuentaHasta VARCHAR(30) = " & lcCuentaContableHasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcAuxDesde VARCHAR(30) = " & lcAuxiliarDesde)
            loComandoSeleccionar.AppendLine("DECLARE @lcAuxHasta VARCHAR(30) = " & lcAuxiliarHasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lcActivo VARCHAR(30) = " & lcActivo)
            loComandoSeleccionar.AppendLine("DECLARE @lcPasivo VARCHAR(30) = " & lcPasivo)
            loComandoSeleccionar.AppendLine("DECLARE @lcCapital VARCHAR(30) = " & lcCapital)
            loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10) = 0")
            loComandoSeleccionar.AppendLine("")
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
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < " & lcCorte00 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS saldo,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte00 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte01 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto01,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte01 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte02 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto02,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte02 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte03 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto03,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte03 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte04 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto04,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte04 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte05 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END") 
			loComandoSeleccionar.AppendLine("		) AS Monto05,") 
			loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte05 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte06 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero") 
			loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Monto06,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte06 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte07 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Monto07,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte07 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte08 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Monto08,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte08 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte09 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Monto09,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte09 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte10 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Monto10,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte10 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte11 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Monto11,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>=" & lcCorte11 & " AND Renglones_Comprobantes.Fec_Ini<" & lcCorte12 & ")")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Monto12")
            loComandoSeleccionar.AppendLine("INTO	#tmpMovimientos") 
			loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables AS CC") 
			loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ") 
			loComandoSeleccionar.AppendLine("		INNER JOIN Comprobantes") 
			loComandoSeleccionar.AppendLine("			ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ") 
			loComandoSeleccionar.AppendLine("				AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("				AND Comprobantes.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			)")
			loComandoSeleccionar.AppendLine("		ON CC.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")
            loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.fec_ini < " & lcCorte12 & ")")
            loComandoSeleccionar.AppendLine("WHERE	CC.Movimiento=1")
            loComandoSeleccionar.AppendLine("	AND	CC.Cod_Cue						BETWEEN @lcCuentaDesde	AND	@lcCuentaHasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Cod_Aux	BETWEEN @lcAuxDesde	AND	@lcAuxHasta")
            loComandoSeleccionar.AppendLine("GROUP BY CC.Cod_Cue")
            loComandoSeleccionar.AppendLine("")
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
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Saldo)							AS Saldo_Inicial,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto01)	                    AS Monto_01,") 
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto02)	                    AS Monto_02,") 
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto03)	                    AS Monto_03,")  
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto04)	                    AS Monto_04,")
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto05)	                    AS Monto_05,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto06)	                    AS Monto_06,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto07)	                    AS Monto_07,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto08)	                    AS Monto_08,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto09)	                    AS Monto_09,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto10)	                    AS Monto_10,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto11)	                    AS Monto_11,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto12)	                    AS Monto_12,")
            loComandoSeleccionar.AppendLine("			(SUM(#tmpMovimientos.Monto01)+SUM(#tmpMovimientos.Monto02)+SUM(#tmpMovimientos.Monto03)+")
            loComandoSeleccionar.AppendLine("			 SUM(#tmpMovimientos.Monto04)+SUM(#tmpMovimientos.Monto05)+SUM(#tmpMovimientos.Monto06)+")
            loComandoSeleccionar.AppendLine("			 SUM(#tmpMovimientos.Monto07)+SUM(#tmpMovimientos.Monto08)+SUM(#tmpMovimientos.Monto09)+")
            loComandoSeleccionar.AppendLine("			 SUM(#tmpMovimientos.Monto10)+SUM(#tmpMovimientos.Monto11)+SUM(#tmpMovimientos.Monto12)) AS Acumulado--,")
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
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_5 ")
            If llSoloMovimientos Then
                loComandoSeleccionar.AppendLine("HAVING     ABS(SUM(#tmpMovimientos.Monto01))+ABS(SUM(#tmpMovimientos.Monto02))+ABS(SUM(#tmpMovimientos.Monto03))")
                loComandoSeleccionar.AppendLine("           + ABS(SUM(#tmpMovimientos.Monto04))+ABS(SUM(#tmpMovimientos.Monto05))+ABS(SUM(#tmpMovimientos.Monto06))")
                loComandoSeleccionar.AppendLine("		    + ABS(SUM(#tmpMovimientos.Monto07))+ABS(SUM(#tmpMovimientos.Monto08))+ABS(SUM(#tmpMovimientos.Monto09))")
                loComandoSeleccionar.AppendLine("           + ABS(SUM(#tmpMovimientos.Monto10))+ABS(SUM(#tmpMovimientos.Monto11))+ABS(SUM(#tmpMovimientos.Monto12)) > 0")

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

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rEGanancias_Perdidas_Mensual", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvTRF_rEGanancias_Perdidas_Mensual.ReportSource = loObjetoReporte

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
