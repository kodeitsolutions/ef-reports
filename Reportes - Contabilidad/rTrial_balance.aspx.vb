'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'Imports Microsoft.Office.Interop

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTrial_balance"
'-------------------------------------------------------------------------------------------'
Partial Class rTrial_balance

	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
	'Dim loAppExcel As Excel.Application

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try

			Dim lcFechaDesde			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcFechaHasta			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcCuentaContableDesde	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
			Dim lcCuentaContableHasta	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
			Dim lcCentroCostoDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcCentroCostoHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcCuentaGastoDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcCuentaGastoHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			Dim lcAuxiliarDesde 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcAuxiliarHasta 		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcMonedaDesde			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcMonedaHasta			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
			
			'Dim lnValor001 As Integer = goOpciones.mObtener("ANCNIV1", "N")
			'Dim lnValor002 As Integer = goOpciones.mObtener("ANCNIV2", "N")
			'Dim lnValor003 As Integer = goOpciones.mObtener("ANCNIV3", "N")
			'Dim lnValor004 As Integer = goOpciones.mObtener("ANCNIV4", "N")
			'Dim lnValor005 As Integer = goOpciones.mObtener("ANCNIV5", "N")
			'Dim lnValor006 As Integer = goOpciones.mObtener("ANCNIV6", "N")
			'Dim lnLongitud001 As Integer = lnValor001
			'Dim lnLongitud002 As Integer = (lnValor001 + lnValor002)
			'Dim lnLongitud003 As Integer = (lnValor001 + lnValor002 + lnValor003)
			'Dim lnLongitud004 As Integer = (lnValor001 + lnValor002 + lnValor003 + lnValor004)
			'Dim lnLongitud005 As Integer = (lnValor001 + lnValor002 + lnValor003 + lnValor004 + lnValor005)
			'Dim lnLongitud006 As Integer = (lnValor001 + lnValor002 + lnValor003 + lnValor004 + lnValor005 + lnValor006)
			'Dim lnNiveles As Integer = goOpciones.mObtener("CANNIVCUE", "N")

			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()
			Dim lnTamaño1 As Integer = CInt(goOpciones.mObtener("ANCNIV1", "")) 
			Dim lnTamaño2 As Integer = CInt(goOpciones.mObtener("ANCNIV2", "")) 
			Dim lnTamaño3 As Integer = CInt(goOpciones.mObtener("ANCNIV3", "")) 
			Dim lnTamaño4 As Integer = CInt(goOpciones.mObtener("ANCNIV4", "")) 
			Dim lnTamaño5 As Integer = CInt(goOpciones.mObtener("ANCNIV5", "")) 
			Dim lnTamaño6 As Integer = CInt(goOpciones.mObtener("ANCNIV6", ""))
			
			Dim lcSeparador As String = CStr(goOpciones.mObtener("CARSEPCUE", "")).Trim()		'Separador de Niveles de la empresa

			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE	@lnCero DECIMAL(28, 10);")
			loComandoSeleccionar.AppendLine("SET		@lnCero = 0;")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño1 INT; SET @lnTamaño1= " & lnTamaño1.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño2 INT; SET @lnTamaño2= " & lnTamaño2.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño3 INT; SET @lnTamaño3= " & lnTamaño3.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño4 INT; SET @lnTamaño4= " & lnTamaño4.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño5 INT; SET @lnTamaño5= " & lnTamaño5.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @lnTamaño6 INT; SET @lnTamaño6= " & lnTamaño6.ToString() & ";")
			loComandoSeleccionar.AppendLine("DECLARE @llUsarTasa BIT; SET @llUsarTasa = 0;")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("SELECT		CC.Cod_Cue								AS Cod_Cue,") 
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
			loComandoSeleccionar.AppendLine("		) AS Monto		") 
			loComandoSeleccionar.AppendLine("INTO	#tmpMovimientos") 
			loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables AS CC") 
			loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ") 
			loComandoSeleccionar.AppendLine("		INNER JOIN Comprobantes") 
			loComandoSeleccionar.AppendLine("			ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ") 
			loComandoSeleccionar.AppendLine("				AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
			loComandoSeleccionar.AppendLine("				AND Comprobantes.Tipo = 'Diario' AND Comprobantes.Status <> 'Anulado'") 
			loComandoSeleccionar.AppendLine("			)")
			loComandoSeleccionar.AppendLine("		ON CC.Cod_Cue = Renglones_Comprobantes.Cod_Cue ") 
			loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.fec_ini <= " & lcFechaHasta & ")") 
			loComandoSeleccionar.AppendLine("WHERE	CC.Movimiento=1") 
			loComandoSeleccionar.AppendLine("		AND	CC.Cod_Cue						BETWEEN " & lcCuentaContableDesde	& "	AND	" & lcCuentaContableHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Cen	BETWEEN " & lcCentroCostoDesde		& "	AND	" & lcCentroCostoHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Gas	BETWEEN " & lcCuentaGastoDesde		& "	AND	" & lcCuentaGastoHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Aux	BETWEEN " & lcAuxiliarDesde			& "	AND	" & lcAuxiliarHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Mon	BETWEEN " & lcMonedaDesde			& "	AND	" & lcMonedaHasta)
			loComandoSeleccionar.AppendLine("GROUP BY CC.Cod_Cue ") 
			loComandoSeleccionar.AppendLine("") 
			
			If (lcSeparador <> "") Then
				Dim lcSeparadorSQL  As String = goServicios.mObtenerCampoFormatoSQL("%" & lcSeparador)
				loComandoSeleccionar.AppendLine("") 
				loComandoSeleccionar.AppendLine("UPDATE	#tmpMovimientos") 
				loComandoSeleccionar.AppendLine("SET		Cod_Cue = SUBSTRING(Cod_Cue, 1, LEN(Cod_Cue) - 1)") 
				loComandoSeleccionar.AppendLine("WHERE	ISNULL(Cod_Cue, '') LIKE " & lcSeparadorSQL ) 
				loComandoSeleccionar.AppendLine("") 
				loComandoSeleccionar.AppendLine("") 
			End If
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("SELECT		Cuentas_Contables.Cod_Cue							AS Cod_Cue,")
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Nom_Cue							AS Nom_Cue,")
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Saldo)							AS Saldo_Inicial,") 
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Debe)							AS Debe,") 
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Haber)							AS Haber,") 
			loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto + #tmpMovimientos.Saldo)	AS Saldo_Actual") 
			loComandoSeleccionar.AppendLine("FROM		#tmpMovimientos") 
			loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Contables") 
			loComandoSeleccionar.AppendLine("		ON	Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Cue") 
			loComandoSeleccionar.AppendLine("GROUP BY	Cuentas_Contables.Cod_Cue, ") 
			loComandoSeleccionar.AppendLine("			Cuentas_Contables.Nom_Cue ") 
			loComandoSeleccionar.AppendLine("HAVING	ABS(SUM(#tmpMovimientos.Saldo)) + ABS(SUM(#tmpMovimientos.Debe)) + ABS(SUM(#tmpMovimientos.Haber)) > 0") 
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento) 
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos") 
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")

			loComandoSeleccionar.AppendLine("")

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

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTrial_balance", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrTrial_balance.ReportSource = loObjetoReporte

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
' MAT: 28/04/11: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' RJG: 22/06/12: Optimización de velocidad: Se rehizo el SELECT y el RPT para mayor			'
'				 velocidad.																	'
'-------------------------------------------------------------------------------------------' 
