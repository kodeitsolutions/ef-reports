'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CCMV_rMayor_Analitico_Parcial"
'-------------------------------------------------------------------------------------------'
Partial Class CCMV_rMayor_Analitico_Parcial

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
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lnValor001 As Integer = goOpciones.mObtener("ANCNIV1", "N")
            Dim lnValor002 As Integer = goOpciones.mObtener("ANCNIV2", "N")
            Dim lnValor003 As Integer = goOpciones.mObtener("ANCNIV3", "N")
            Dim lnValor004 As Integer = goOpciones.mObtener("ANCNIV4", "N")
            Dim lnValor005 As Integer = goOpciones.mObtener("ANCNIV5", "N")
            Dim lnValor006 As Integer = goOpciones.mObtener("ANCNIV6", "N")
            Dim lnLongitud001 As Integer = lnValor001
            Dim lnLongitud002 As Integer = (lnValor001 + lnValor002)
            Dim lnLongitud003 As Integer = (lnValor001 + lnValor002 + lnValor003)
            Dim lnLongitud004 As Integer = (lnValor001 + lnValor002 + lnValor003 + lnValor004)
            Dim lnLongitud005 As Integer = (lnValor001 + lnValor002 + lnValor003 + lnValor004 + lnValor005)
            Dim lnLongitud006 As Integer = (lnValor001 + lnValor002 + lnValor003 + lnValor004 + lnValor005 + lnValor006)
            Dim lnNiveles As Integer = goOpciones.mObtener("CANNIVCUE", "N")

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE	@lcFechaInicio		DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@lcFechaFin			DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@lcCuentaInicio 	VARCHAR(100)")
            loComandoSeleccionar.AppendLine("DECLARE	@lcCuentaFin		VARCHAR(100)")
            loComandoSeleccionar.AppendLine("DECLARE	@lcCostoInicio		VARCHAR(100)")
            loComandoSeleccionar.AppendLine("DECLARE	@lcCostoFin			VARCHAR(100)")
            loComandoSeleccionar.AppendLine("DECLARE	@lcGastoInicio		VARCHAR(100)")
            loComandoSeleccionar.AppendLine("DECLARE	@lcGastoFin			VARCHAR(100)")
            loComandoSeleccionar.AppendLine("DECLARE	@lcAuxiliarInicio	VARCHAR(100)")
            loComandoSeleccionar.AppendLine("DECLARE	@lcAuxiliarFin		VARCHAR(100)")
            loComandoSeleccionar.AppendLine("DECLARE	@lcMonedaInicio		VARCHAR(100)")
            loComandoSeleccionar.AppendLine("DECLARE	@lcMonedaFin		VARCHAR(100)")
            loComandoSeleccionar.AppendLine("DECLARE	@lcDocumentoInicio	VARCHAR(100)")
            loComandoSeleccionar.AppendLine("DECLARE	@lcDocumentoFin		VARCHAR(100)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET	@lcFechaInicio		= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("SET	@lcFechaFin			= " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("SET	@lcCuentaInicio		= " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("SET	@lcCuentaFin		= " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("SET	@lcCostoInicio		= " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("SET	@lcCostoFin			= " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("SET	@lcGastoInicio		= " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("SET	@lcGastoFin			= " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("SET	@lcAuxiliarInicio	= " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("SET	@lcAuxiliarFin		= " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("SET	@lcMonedaInicio		= " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("SET	@lcMonedaFin		= " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("SET	@lcDocumentoInicio	= " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("SET	@lcDocumentoFin		= " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE	@lnCero DECIMAL(28,10)")
            loComandoSeleccionar.AppendLine("SET		@lnCero = 0")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Cuentas_Contables.Cod_Cue,")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Nom_Cue,")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab)	AS Saldo_Inicial,")
            'loComandoSeleccionar.AppendLine("			SUM(ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)) AS Saldo_Inicial,")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas_Contables.Cod_Cue,1, " & lnLongitud001 & ")	AS	Nivel01, ")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas_Contables.Cod_Cue,1, " & lnLongitud002 & ")	AS	Nivel02, ")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas_Contables.Cod_Cue,1, " & lnLongitud003 & ")	AS	Nivel03, ")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas_Contables.Cod_Cue,1, " & lnLongitud004 & ")	AS	Nivel04, ")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas_Contables.Cod_Cue,1, " & lnLongitud005 & ")	AS	Nivel05, ")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas_Contables.Cod_Cue,1, " & lnLongitud006 & ")	AS	Nivel06	 ")
            loComandoSeleccionar.AppendLine("INTO 		#tmpSaldosIniciales")
            loComandoSeleccionar.AppendLine("FROM 		Comprobantes")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Comprobantes ")
            loComandoSeleccionar.AppendLine("		ON	Comprobantes.Documento = Renglones_Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("		AND Comprobantes.Adicional = Renglones_Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine("		AND Comprobantes.Documento				BETWEEN @lcDocumentoInicio AND @lcDocumentoFin")
            loComandoSeleccionar.AppendLine("		AND	Comprobantes.Status					<> 'Anulado' ")
            loComandoSeleccionar.AppendLine("		AND	Comprobantes.Tipo					=	'Diario' ")
            loComandoSeleccionar.AppendLine("		AND	Renglones_Comprobantes.Fec_Ini		< @lcFechaInicio")
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Cue		BETWEEN @lcCuentaInicio AND @lcCuentaFin")
            loComandoSeleccionar.AppendLine("       AND Renglones_Comprobantes.Cod_Cen		BETWEEN @lcCostoInicio AND @lcCostoFin ")
            loComandoSeleccionar.AppendLine("       AND Renglones_Comprobantes.Cod_Gas		BETWEEN @lcGastoInicio AND @lcGastoFin ")
            loComandoSeleccionar.AppendLine("       AND Renglones_Comprobantes.Cod_Aux		BETWEEN @lcAuxiliarInicio AND @lcAuxiliarFin ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Mon		BETWEEN @lcMonedaInicio AND @lcMonedaFin ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Documento	BETWEEN @lcDocumentoInicio AND @lcDocumentoFin ")
            loComandoSeleccionar.AppendLine("	RIGHT JOIN Cuentas_Contables")
            loComandoSeleccionar.AppendLine("        ON	Renglones_Comprobantes.Cod_Cue = Cuentas_Contables.Cod_Cue        ")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Contables.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin")
            loComandoSeleccionar.AppendLine("GROUP BY	Cuentas_Contables.Cod_Cue, Cuentas_Contables.Nom_Cue ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		ROW_NUMBER() ")
            loComandoSeleccionar.AppendLine("				OVER (ORDER BY	#tmpSaldosIniciales.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("								Renglones_Comprobantes.Documento, ")
            loComandoSeleccionar.AppendLine("								Renglones_Comprobantes.Renglon) ")
            loComandoSeleccionar.AppendLine("															AS Posicion,")
            loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Documento				AS Documento, ")
            loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Renglon					AS Renglon, ")
            loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Comentario				AS Comentario, ")
            loComandoSeleccionar.AppendLine("			#tmpSaldosIniciales.Cod_Cue						AS Cod_Cue, ")
            loComandoSeleccionar.AppendLine("			#tmpSaldosIniciales.Nom_Cue						AS Nom_Cue, ")
            loComandoSeleccionar.AppendLine("			#tmpSaldosIniciales.Nivel01						AS Nivel01, ")
            loComandoSeleccionar.AppendLine("			#tmpSaldosIniciales.Nivel02						AS Nivel02, ")
            loComandoSeleccionar.AppendLine("			#tmpSaldosIniciales.Nivel03						AS Nivel03, ")
            loComandoSeleccionar.AppendLine("			#tmpSaldosIniciales.Nivel04						AS Nivel04, ")
            loComandoSeleccionar.AppendLine("			#tmpSaldosIniciales.Nivel05						AS Nivel05, ")
            loComandoSeleccionar.AppendLine("			#tmpSaldosIniciales.Nivel06						AS Nivel06, ")
            loComandoSeleccionar.AppendLine("			#tmpSaldosIniciales.Saldo_Inicial				AS Saldo_Inicial, ")
            loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Cod_Cen					AS Cod_Cen, ")
            loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Cod_Gas					AS Cod_Gas, ")
            loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Cod_Aux					AS Cod_Aux, ")
            loComandoSeleccionar.AppendLine("			ISNULL(Auxiliares.Nom_Aux, '')					AS Nom_Aux, ")
            loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Fec_Ini					AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Mon_Deb					AS Mon_Deb,")
            loComandoSeleccionar.AppendLine("			Renglones_Comprobantes.Mon_Hab					AS Mon_Hab,")
            loComandoSeleccionar.AppendLine("			@lnCero											AS Monto ")
            loComandoSeleccionar.AppendLine("INTO		#tmpComprobantes ")
            loComandoSeleccionar.AppendLine("FROM		Comprobantes")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Comprobantes		")
            loComandoSeleccionar.AppendLine("		ON	Renglones_Comprobantes.Documento	= Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine("		AND Comprobantes.Documento				BETWEEN @lcDocumentoInicio AND @lcDocumentoFin")
            loComandoSeleccionar.AppendLine("		AND	Comprobantes.Status					<>	'Anulado' ")
            loComandoSeleccionar.AppendLine("		AND	Comprobantes.Tipo					=	'Diario' ")
            loComandoSeleccionar.AppendLine("		AND	Renglones_Comprobantes.Fec_Ini		BETWEEN @lcFechaInicio AND @lcFechaFin ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Cue		BETWEEN @lcCuentaInicio AND @lcCuentaFin")
            loComandoSeleccionar.AppendLine("       	AND Renglones_Comprobantes.Cod_Cen		BETWEEN @lcCostoInicio AND @lcCostoFin ")
            loComandoSeleccionar.AppendLine("       	AND Renglones_Comprobantes.Cod_Gas		BETWEEN @lcGastoInicio AND @lcGastoFin ")
            loComandoSeleccionar.AppendLine("       	AND Renglones_Comprobantes.Cod_Aux		BETWEEN @lcAuxiliarInicio AND @lcAuxiliarFin ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Mon		BETWEEN @lcMonedaInicio AND @lcMonedaFin 		")
            loComandoSeleccionar.AppendLine("	JOIN #tmpSaldosIniciales")
            loComandoSeleccionar.AppendLine("		ON	Renglones_Comprobantes.Cod_Cue = #tmpSaldosIniciales.Cod_Cue    ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Auxiliares")
            loComandoSeleccionar.AppendLine("        ON	Renglones_Comprobantes.Cod_Aux = Auxiliares.Cod_Aux        ")
            loComandoSeleccionar.AppendLine("ORDER BY #tmpSaldosIniciales.Cod_Cue, Renglones_Comprobantes.Documento, Renglones_Comprobantes.Renglon  ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpSaldosIniciales")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE	#tmpComprobantes")
            loComandoSeleccionar.AppendLine("SET		Monto = (	SELECT TOP	1 ")
            loComandoSeleccionar.AppendLine("								SUM(Mon_Deb - Mon_Hab)+MIN(Saldo_Inicial) ")
            loComandoSeleccionar.AppendLine("					FROM		#tmpComprobantes AS L")
            loComandoSeleccionar.AppendLine("					WHERE		L.Cod_Cue = #tmpComprobantes.Cod_Cue")
            loComandoSeleccionar.AppendLine("						AND		L.Posicion <= #tmpComprobantes.Posicion")
            loComandoSeleccionar.AppendLine("				)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	* ")
            loComandoSeleccionar.AppendLine("FROM	#tmpComprobantes       ")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Cue, Documento, Renglon  ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpComprobantes       ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes", 450)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CCMV_rMayor_Analitico_Parcial", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCCMV_rMayor_Analitico_Parcial.ReportSource = loObjetoReporte

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
' GSC: 27/06/12: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' RJG: 16/12/13: Se corrigió el el cálculo del saldo inicial de las cuentas.                '
'-------------------------------------------------------------------------------------------' 
