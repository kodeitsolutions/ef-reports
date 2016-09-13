'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Diario"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Diario

	Inherits vis2formularios.frmReporte

	Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
	
	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try

			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
			Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
			Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
			Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
			Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
			Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
			Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
			Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
			Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
			Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
			Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
			
 			Dim llSoloMovimientos As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(6)).Trim().ToUpper().Equals("SI")
 			Dim llMostrarCuentasRaiz As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(8)).Trim().ToUpper().Equals("SI")
		
			Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()

			loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Desde VARCHAR(100)	")	
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro1Hasta VARCHAR(100)	")	
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Desde DATETIME	")		
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro2Hasta DATETIME	")		
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Desde VARCHAR(100)	")	
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro3Hasta VARCHAR(100)	")	
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Desde VARCHAR(100)	")	
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro4Hasta VARCHAR(100)	")	
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Desde VARCHAR(100)	")	
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro5Hasta VARCHAR(100)	")	
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro6Desde VARCHAR(100)	")	
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro7Desde VARCHAR(100)	")	
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro8Desde VARCHAR(100)	")	
			loComandoSeleccionar.AppendLine("DECLARE @lcParametro8Hasta VARCHAR(100)	")
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
			loComandoSeleccionar.AppendLine("SET		@lcParametro7Desde = " & lcParametro6Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro8Desde = " & lcParametro7Desde)
			loComandoSeleccionar.AppendLine("SET		@lcParametro8Hasta = " & lcParametro7Hasta)
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
			loComandoSeleccionar.AppendLine("			@llFalso						AS Nivel_7")
			loComandoSeleccionar.AppendLine("INTO		#tmpCuentas")
			loComandoSeleccionar.AppendLine("FROM		Cuentas_Contables")
			loComandoSeleccionar.AppendLine("WHERE		Cod_Cue BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
			If Not llMostrarCuentasRaiz Then
				loComandoSeleccionar.AppendLine("	AND		movimiento = 1")
			End If
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
			loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_7								AS Nivel_7")
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
			loComandoSeleccionar.AppendLine("						AND Comprobantes.Tipo = @lcParametro6Desde")
			loComandoSeleccionar.AppendLine("					AND Comprobantes.Status <> 'Anulado'")
			loComandoSeleccionar.AppendLine("				WHERE		Renglones_Comprobantes.Fec_Ini	BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Gas	BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Cen	BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Aux	BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine("						AND Renglones_Comprobantes.Cod_Mon	BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
			loComandoSeleccionar.AppendLine("			) AS tmpRenglones ")
			loComandoSeleccionar.AppendLine("		ON tmpRenglones.Cod_Cue = #tmpCuentas.Cod_Cue")
			loComandoSeleccionar.AppendLine("GROUP BY	#tmpCuentas.Nivel, #tmpCuentas.Cod_cue, #tmpCuentas.Nom_Cue, ")
			loComandoSeleccionar.AppendLine("			#tmpCuentas.Movimiento,")
			loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_1, #tmpCuentas.Nivel_2, #tmpCuentas.Nivel_3, ")
			loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_4, #tmpCuentas.Nivel_5, #tmpCuentas.Nivel_6, ")
			loComandoSeleccionar.AppendLine("			#tmpCuentas.Nivel_7")
			loComandoSeleccionar.AppendLine("ORDER BY #tmpCuentas.Cod_cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpCuentas")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("--*********************************************************")
			loComandoSeleccionar.AppendLine("-- Calcula el máximo nivel                                *")
			loComandoSeleccionar.AppendLine("--*********************************************************")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DECLARE @lcMaximoNivel INT")
			loComandoSeleccionar.AppendLine("SET @lcMaximoNivel = (SELECT MAX(Nivel) FROM #tmpTodos)")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			If llSoloMovimientos Then 
				loComandoSeleccionar.AppendLine("--*********************************************************")
				loComandoSeleccionar.AppendLine("-- Elimina los niveles máximos y sin movimientos          *")
				loComandoSeleccionar.AppendLine("--*********************************************************")
				loComandoSeleccionar.AppendLine("DELETE		FROM #tmpTodos")
				loComandoSeleccionar.AppendLine("WHERE		#tmpTodos.Nivel = @lcMaximoNivel")
				loComandoSeleccionar.AppendLine("	AND		(ABS(Mon_Deb)+ABS(Mon_Hab)) = @lnCero")
				loComandoSeleccionar.AppendLine("")
				loComandoSeleccionar.AppendLine("")
				loComandoSeleccionar.AppendLine("")
				loComandoSeleccionar.AppendLine("--*********************************************************")
				loComandoSeleccionar.AppendLine("-- Elimina los niveles superiores y sin movimientos       *")
				loComandoSeleccionar.AppendLine("--*********************************************************")
				loComandoSeleccionar.AppendLine("SELECT	@lnCero As SubTotal, Cod_Cue, Nivel")
				loComandoSeleccionar.AppendLine("INTO	#tmpTotales")
				loComandoSeleccionar.AppendLine("FROM	#tmpTodos")
				loComandoSeleccionar.AppendLine("WHERE	#tmpTodos.Nivel < @lcMaximoNivel")
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
				loComandoSeleccionar.AppendLine("		FROM		(	SELECT		SUM(A.Mon_Deb) - SUM(A.Mon_Hab) AS SubTotal,")		
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
			loComandoSeleccionar.AppendLine("SELECT		@lcMaximoNivel AS Nivel_Maximo, ")
			loComandoSeleccionar.AppendLine("			#tmpTodos.* ")
			loComandoSeleccionar.AppendLine("FROM		#tmpTodos")
			loComandoSeleccionar.AppendLine("ORDER BY	Cod_cue")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpTodos")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			
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

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Diario", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)

			Me.mFormatearCamposReporte(loObjetoReporte)

			Me.crvrLibro_Diario.ReportSource = loObjetoReporte

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
' RJG: 24/10/11: Codigo inicial, a partir de Balance General.								'
'-------------------------------------------------------------------------------------------' 
' RJG: 27/10/11: Corrección en cálculo de saldos inicial y final. Corrección en filtro.	'
'-------------------------------------------------------------------------------------------' 
' RJG: 16/11/11: Agregado filtro de la estructura superior cuando en el detalle no hay		'
'				 movimientos (y el usuario indicó el filtro "Solo Moviminetos = SI".		'
'-------------------------------------------------------------------------------------------' 
' RJG: 06/12/11: Se agregó la igualdad de campo Adicional en las uniones entre Comprobantes	'
'				 y sus renglones.															'
'-------------------------------------------------------------------------------------------' 
' RJG: 12/01/12: Se agregó un parámetro más (llMostrarCuentasRaiz) para inicar si se		'
'				 muestran las cuentas raiz (estructura) o no (tabla "plana").				'
'-------------------------------------------------------------------------------------------' 
' RJG: 22/06/12: Se modificaron los parámetros para usar valores concatenados en lugar de	'
'				 variables SQL (optimización de velocidad).									'
'-------------------------------------------------------------------------------------------' 
