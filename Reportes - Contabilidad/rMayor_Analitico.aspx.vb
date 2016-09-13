'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rMayor_Analitico"
'-------------------------------------------------------------------------------------------'
Partial Class rMayor_Analitico

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
			Dim lcDocumentoDesde		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcDocumentoHasta		As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
					
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()
			
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE	@lnCero DECIMAL(28, 10);")
			loComandoSeleccionar.AppendLine("SET		@lnCero = 0;")
			loComandoSeleccionar.AppendLine("DECLARE @llUsarTasa BIT; SET @llUsarTasa = 0;")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	ROW_NUMBER() ") 
			loComandoSeleccionar.AppendLine("			OVER (PARTITION BY Renglones_Comprobantes.Cod_Cue ") 
			loComandoSeleccionar.AppendLine("				ORDER BY Renglones_Comprobantes.Documento, Renglones_Comprobantes.Renglon)			AS Posicion, ") 
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Documento															AS Documento, ")
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Renglon																AS Renglon, ") 
			loComandoSeleccionar.AppendLine("		CAST(Renglones_Comprobantes.Comentario AS VARCHAR(MAX))										AS Comentario, ") 
			loComandoSeleccionar.AppendLine("		CC.Cod_Cue																					AS Cod_Cue, ")
			loComandoSeleccionar.AppendLine("		CC.Nom_Cue																					AS Nom_Cue, ")
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Cen																AS Cod_Cen, ") 
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Gas																AS Cod_Gas, ") 
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Aux																AS Cod_Aux, ") 
			loComandoSeleccionar.AppendLine("		ISNULL(Auxiliares.Nom_Aux, '')																AS Nom_Aux, ")
			loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Fec_Ini																AS Fec_Ini, ") 
			loComandoSeleccionar.AppendLine("		ISNULL(Iniciales.Saldo_Inicial, @lnCero)													AS Saldo_Inicial,") 
			loComandoSeleccionar.AppendLine("		(ISNULL(Renglones_Comprobantes.Mon_Deb, @lnCero)") 
			loComandoSeleccionar.AppendLine("				*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)")
			loComandoSeleccionar.AppendLine("			)																						AS Debe,") 
			loComandoSeleccionar.AppendLine("		(ISNULL(Renglones_Comprobantes.Mon_Hab, @lnCero)") 
			loComandoSeleccionar.AppendLine("				*(CASE WHEN @llUsarTasa=1 THEN ISNULL(Renglones_Comprobantes.Tasa,1) ELSE 1 END)") 
			loComandoSeleccionar.AppendLine("			)																						AS Haber")
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
			loComandoSeleccionar.AppendLine("		AND (Renglones_Comprobantes.fec_ini BETWEEN " & lcFechaDesde & " AND " & lcFechaHasta & ")")
			loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN (") 
			loComandoSeleccionar.AppendLine("			SELECT	RC.Cod_Cue						AS Cod_Cue,") 
			loComandoSeleccionar.AppendLine("					SUM(RC.Mon_Deb - RC.Mon_Hab)	AS Saldo_Inicial")
			loComandoSeleccionar.AppendLine("			FROM	Renglones_Comprobantes AS RC ") 
			loComandoSeleccionar.AppendLine("				INNER JOIN Comprobantes AS C") 
			loComandoSeleccionar.AppendLine("					ON (RC.Adicional = C.Adicional ") 
			loComandoSeleccionar.AppendLine("						AND RC.Documento = C.Documento")
			loComandoSeleccionar.AppendLine("						AND C.Tipo = 'Diario' AND C.Status <> 'Anulado'") 
			loComandoSeleccionar.AppendLine("						AND C.Documento BETWEEN " & lcDocumentoDesde & " AND " & lcDocumentoHasta & ")")
			loComandoSeleccionar.AppendLine("			WHERE	(RC.fec_ini < " & lcFechaDesde & " )") 
			loComandoSeleccionar.AppendLine("			AND	RC.Cod_Cue	BETWEEN " & lcCuentaContableDesde	& "	AND	" & lcCuentaContableHasta)
			loComandoSeleccionar.AppendLine("			AND RC.Cod_Cen	BETWEEN " & lcCentroCostoDesde		& "	AND	" & lcCentroCostoHasta)
			loComandoSeleccionar.AppendLine("			AND RC.Cod_Gas	BETWEEN " & lcCuentaGastoDesde		& "	AND	" & lcCuentaGastoHasta)
			loComandoSeleccionar.AppendLine("			AND RC.Cod_Aux	BETWEEN " & lcAuxiliarDesde			& "	AND	" & lcAuxiliarHasta)
			loComandoSeleccionar.AppendLine("			AND RC.Cod_Mon	BETWEEN " & lcMonedaDesde			& "	AND	" & lcMonedaHasta)
			loComandoSeleccionar.AppendLine("			GROUP BY RC.Cod_Cue")
			loComandoSeleccionar.AppendLine("			) AS Iniciales") 
			loComandoSeleccionar.AppendLine("		ON CC.Cod_Cue = Iniciales.Cod_Cue ") 
			loComandoSeleccionar.AppendLine("	LEFT JOIN Auxiliares")
			loComandoSeleccionar.AppendLine("       ON	Renglones_Comprobantes.Cod_Aux = Auxiliares.Cod_Aux") 
			loComandoSeleccionar.AppendLine("WHERE	CC.Movimiento = 1") 
			loComandoSeleccionar.AppendLine("		AND	CC.Cod_Cue						BETWEEN " & lcCuentaContableDesde	& "	AND	" & lcCuentaContableHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Cen	BETWEEN " & lcCentroCostoDesde		& "	AND	" & lcCentroCostoHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Gas	BETWEEN " & lcCuentaGastoDesde		& "	AND	" & lcCuentaGastoHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Aux	BETWEEN " & lcAuxiliarDesde			& "	AND	" & lcAuxiliarHasta)
			loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Mon	BETWEEN " & lcMonedaDesde			& "	AND	" & lcMonedaHasta)
			loComandoSeleccionar.AppendLine("ORDER BY CC.Cod_Cue, Posicion")
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("CREATE Clustered INDEX PK_tmpMovimientos_Cuenta_Posicion ON #tmpMovimientos(Cod_Cue, Posicion)")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT	#tmpMovimientos.Posicion 								AS Posicion,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Documento 								AS Documento,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Renglon 								AS Renglon,		")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Comentario								AS Comentario,	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Cue 								AS Cod_Cue, 	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Nom_Cue 								AS Nom_Cue, 	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Cen 								AS Cod_Cen, 	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Gas 								AS Cod_Gas, 	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Cod_Aux 								AS Cod_Aux, 	")
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Nom_Aux 								AS Nom_Aux, 	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Fec_Ini 								AS Fec_Ini, 	") 
			'loComandoSeleccionar.AppendLine("		#tmpMovimientos.Saldo_Inicial							AS Original,	") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Saldo_Inicial							AS Saldo_Inicial,") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Debe									AS Debe,		") 
			loComandoSeleccionar.AppendLine("		#tmpMovimientos.Haber									AS Haber,		") 
			loComandoSeleccionar.AppendLine("		(#tmpMovimientos.Debe - #tmpMovimientos.Haber)			AS Saldo_Actual	") 
			loComandoSeleccionar.AppendLine("FROM	#tmpMovimientos") 
			loComandoSeleccionar.AppendLine("GROUP BY	#tmpMovimientos.Posicion,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Documento,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Renglon,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Comentario,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Cue,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Nom_Cue,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Cen,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Gas,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Aux,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Nom_Aux,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Fec_Ini,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Saldo_Inicial,") 
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Debe,")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Haber")
			loComandoSeleccionar.AppendLine("ORDER BY	#tmpMovimientos.Cod_Cue, ")
			loComandoSeleccionar.AppendLine("			#tmpMovimientos.Posicion") 
			loComandoSeleccionar.AppendLine("") 
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

			Dim loServicios As New cusDatos.goDatos

			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes", 1200)

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

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rMayor_Analitico", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrMayor_Analitico.ReportSource = loObjetoReporte

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
' JJD: 11/10/10: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' JJD: 13/10/10: Ajuste del calculo al monto del campo: Saldo_Actual						'
'-------------------------------------------------------------------------------------------'
' MAT: 16/05/11: Mejora de la vista de diseño												'
'-------------------------------------------------------------------------------------------' 
' RJG: 08/11/11: Ajuste en filtro y uniones. El filtro de fechas ahora aplica a la fecha de	'
'				 los asientos, no a la del comprobante.										'
'-------------------------------------------------------------------------------------------' 
' RJG: 18/11/11: Agregado el Nombre del Auxiliar al Reporte.								'
'-------------------------------------------------------------------------------------------' 
' RJG: 05/12/11: Se agregó la igualdad de campo Adicional en las uniones entre Comprobantes	'
'				 y sus renglones.															'
'-------------------------------------------------------------------------------------------' 
' RJG: 09/02/12: Cambiado el filtro de comprobantes "='Pendiente'" por "<>'Anulado'".		'
'-------------------------------------------------------------------------------------------' 
' RJG: 16/03/12: Ajuste en filtro: Se cambió nuevamente el filtro de fechas para aplicar a	'
'				 la fecha del comprobante en lugar de la fecha del asiento (ver 08/11/11).	'
'-------------------------------------------------------------------------------------------' 
' RJG: 16/03/12: Ajuste en totales: Se pasó el cálculo de los totales acumulados al RPT para'
'				 mejorar el tiempo de la consulta.											'
'-------------------------------------------------------------------------------------------' 
' RJG: 16/12/13: Se corrigieron los filtros en el cálculo del saldo inicial de las cuentas. '
'-------------------------------------------------------------------------------------------' 
