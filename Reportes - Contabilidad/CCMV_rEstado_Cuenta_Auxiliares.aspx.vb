'-------------------------------------------------------------------------------------------'
' Inicio del codigo																			'
'-------------------------------------------------------------------------------------------'
' Importando librerias																		'
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CCMV_rEstado_Cuenta_Auxiliares"											'
'-------------------------------------------------------------------------------------------'
Partial Class CCMV_rEstado_Cuenta_Auxiliares

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcFechaDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcFechaHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcCuentaContableDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcCuentaContableHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcCentroCostoDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcCentroCostoHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcCuentaGastoDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcCuentaGastoHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcAuxiliarDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcAuxiliarHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcMonedaDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcMonedaHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcTipoComprobante As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            'Dim lcSoloConMovimiento 	As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            'Dim lcNivelMostrar			As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))

            Dim llSoloMovimientos As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(7)).Trim().ToUpper().Equals("SI")

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
            Dim lnTamaño1 As Integer = CInt(goOpciones.mObtener("ANCNIV1", ""))
            Dim lnTamaño2 As Integer = CInt(goOpciones.mObtener("ANCNIV2", ""))
            Dim lnTamaño3 As Integer = CInt(goOpciones.mObtener("ANCNIV3", ""))
            Dim lnTamaño4 As Integer = CInt(goOpciones.mObtener("ANCNIV4", ""))
            Dim lnTamaño5 As Integer = CInt(goOpciones.mObtener("ANCNIV5", ""))
            Dim lnTamaño6 As Integer = CInt(goOpciones.mObtener("ANCNIV6", ""))

            Dim lnCantidad As Integer = CInt(goOpciones.mObtener("CANNIVCUE", ""))              'Niveles usados por la empresa
            Dim lnNivelMax As Integer = CInt(cusAplicacion.goReportes.paParametrosIniciales(8)) 'Nivel máximo a mostrar	en reporte
            Dim lcSeparador As String = CStr(goOpciones.mObtener("CARSEPCUE", "")).Trim()       'Separador de Niveles de la empresa

            Dim lnLongMax As Integer = 0

            'Establece el nivel máximo a mostrar
            If lnCantidad <= 0 Then
                lnCantidad = 1
            ElseIf (lnCantidad > 6) Then
                lnCantidad = 6
            End If
            lnCantidad = Math.Min(lnCantidad, lnNivelMax)

            'Establece las longitudes de todos los niveles

            If (lnCantidad < 1) Then
                lnTamaño1 = 0
            ElseIf (lnTamaño1 <= 0) Then
                lnTamaño1 = Math.Max(lnTamaño1, 1)
            End If

            If (lnCantidad < 2) Then
                lnTamaño2 = 0
            ElseIf (lnTamaño2 <= 0) Then
                lnTamaño2 = Math.Max(lnTamaño2, 1)
            End If
            lnTamaño2 += lnTamaño1

            If (lnCantidad < 3) Then
                lnTamaño3 = 0
            ElseIf (lnTamaño3 <= 0) Then
                lnTamaño3 = Math.Max(lnTamaño3, 1)
            End If
            lnTamaño3 += lnTamaño2

            If (lnCantidad < 4) Then
                lnTamaño4 = 0
            ElseIf (lnTamaño4 <= 0) Then
                lnTamaño4 = Math.Max(lnTamaño4, 1)
            End If
            lnTamaño4 += lnTamaño3

            If (lnCantidad < 5) Then
                lnTamaño5 = 0
            ElseIf (lnTamaño5 <= 0) Then
                lnTamaño5 = Math.Max(lnTamaño5, 1)
            End If
            lnTamaño5 += lnTamaño4

            If (lnCantidad < 6) Then
                lnTamaño6 = 0
            ElseIf (lnTamaño6 <= 0) Then
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
                Case 7  'Mostrar Todos
                    lnLongMax = lnTamaño6
            End Select

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
            loComandoSeleccionar.AppendLine("		SUBSTRING(CC.Cod_Cue, 1, CASE WHEN (LEN(RTRIM(CC.Cod_Cue))>@lnTamaño6 AND @lnNivelMax > 6) THEN @lnTamaño6 ELSE 0 END)	AS Cod_Niv_6 ")
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
            loComandoSeleccionar.AppendLine("		AND	CC.Cod_Cue						BETWEEN " & lcCuentaContableDesde & "	AND	" & lcCuentaContableHasta)
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Cen	BETWEEN " & lcCentroCostoDesde & "	AND	" & lcCentroCostoHasta)
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Gas	BETWEEN " & lcCuentaGastoDesde & "	AND	" & lcCuentaGastoHasta)
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Aux	BETWEEN " & lcAuxiliarDesde & "	AND	" & lcAuxiliarHasta)
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Mon	BETWEEN " & lcMonedaDesde & "	AND	" & lcMonedaHasta)
            loComandoSeleccionar.AppendLine("GROUP BY CC.Cod_Cue ")
            loComandoSeleccionar.AppendLine("")



            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  SUM(Inicial) AS Inicial, Cod_Aux, Cod_Cue")
            loComandoSeleccionar.AppendLine("INTO	    #tmpSaldoInicial")
            loComandoSeleccionar.AppendLine("FROM    (")
            loComandoSeleccionar.AppendLine("        SELECT	    0   AS Inicial, ")
            loComandoSeleccionar.AppendLine("        		    ''  AS Cod_Aux, ")
            loComandoSeleccionar.AppendLine("        		    ''  AS Cod_Cue ")
            loComandoSeleccionar.AppendLine("        UNION ALL")
            loComandoSeleccionar.AppendLine("        SELECT	    SUM(Renglones_comprobantes.Mon_Deb - Renglones_comprobantes.Mon_Hab) AS Inicial,")
            loComandoSeleccionar.AppendLine("        		    Renglones_comprobantes.Cod_Aux                                       AS Cod_Aux,")
            loComandoSeleccionar.AppendLine("        		    Renglones_comprobantes.Cod_Cue                                       AS Cod_Cue ")
            loComandoSeleccionar.AppendLine("        FROM	    Comprobantes")
            loComandoSeleccionar.AppendLine("        JOIN        Renglones_Comprobantes ")
            loComandoSeleccionar.AppendLine("        	ON      Renglones_Comprobantes.Documento = Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("        		AND Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine("                AND Comprobantes.Tipo = 'Diario' ")
            loComandoSeleccionar.AppendLine("                AND Comprobantes.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("        WHERE	    Renglones_Comprobantes.Fec_Ini  < " & lcFechaDesde)
            loComandoSeleccionar.AppendLine("        		AND	Renglones_Comprobantes.Cod_Cue	BETWEEN " & lcCuentaContableDesde & " AND " & lcCuentaContableHasta)
            loComandoSeleccionar.AppendLine("        		AND Renglones_Comprobantes.Cod_Cen	BETWEEN " & lcCentroCostoDesde & " AND " & lcCentroCostoHasta)
            loComandoSeleccionar.AppendLine("        		AND Renglones_Comprobantes.Cod_Gas	BETWEEN " & lcCuentaGastoDesde & " AND " & lcCuentaGastoHasta)
            loComandoSeleccionar.AppendLine("        		AND Renglones_Comprobantes.Cod_Aux	BETWEEN " & lcAuxiliarDesde & "	AND " & lcAuxiliarHasta)
            loComandoSeleccionar.AppendLine("        		AND Renglones_Comprobantes.Cod_Mon	BETWEEN " & lcMonedaDesde & " AND " & lcMonedaHasta)
            loComandoSeleccionar.AppendLine("        GROUP  BY   Renglones_comprobantes.Cod_Aux, Renglones_comprobantes.Cod_Cue")
            loComandoSeleccionar.AppendLine("       ) AS X")
            loComandoSeleccionar.AppendLine("GROUP BY Cod_Aux, Cod_Cue")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT	    Comprobantes.Documento, ")
            loComandoSeleccionar.AppendLine("		    Renglones_comprobantes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("		    Renglones_comprobantes.Renglon, ")
            loComandoSeleccionar.AppendLine("		    Renglones_comprobantes.Comentario, ")
            loComandoSeleccionar.AppendLine("		    Renglones_comprobantes.Mon_Deb AS Debe, ")
            loComandoSeleccionar.AppendLine("		    Renglones_comprobantes.Mon_Hab AS Haber, ")
            loComandoSeleccionar.AppendLine("		    Renglones_comprobantes.Cod_Aux, ")
            loComandoSeleccionar.AppendLine("		    Renglones_comprobantes.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("		    Renglones_comprobantes.Nom_Cue, ")

            loComandoSeleccionar.AppendLine("		    ROW_NUMBER()")
            loComandoSeleccionar.AppendLine("		        OVER (  PARTITION BY Renglones_Comprobantes.Cod_Aux ")
            loComandoSeleccionar.AppendLine("		                ORDER BY     Renglones_Comprobantes.Cod_Aux, Renglones_Comprobantes.Fec_Ini")
            loComandoSeleccionar.AppendLine("               ) AS Posicion, ")

            loComandoSeleccionar.AppendLine("		    COALESCE(#tmpSaldoInicial.Inicial, 0) AS Saldo_Inicial, ")
            loComandoSeleccionar.AppendLine("		    #tmpMovimientos.Cod_Niv_1 AS Cod_Niv_1, ")
            loComandoSeleccionar.AppendLine("		    #tmpMovimientos.Cod_Niv_2 AS Cod_Niv_2, ")
            loComandoSeleccionar.AppendLine("		    #tmpMovimientos.Cod_Niv_3 AS Cod_Niv_3, ")
            loComandoSeleccionar.AppendLine("		    #tmpMovimientos.Cod_Niv_4 AS Cod_Niv_4, ")
            loComandoSeleccionar.AppendLine("		    #tmpMovimientos.Cod_Niv_5 AS Cod_Niv_5, ")
            loComandoSeleccionar.AppendLine("		    #tmpMovimientos.Cod_Niv_6 AS Cod_Niv_6 ")
            loComandoSeleccionar.AppendLine("INTO	    #Saldos")
            loComandoSeleccionar.AppendLine("FROM	    Comprobantes")
            loComandoSeleccionar.AppendLine("JOIN       Renglones_Comprobantes ")
            loComandoSeleccionar.AppendLine("			ON (Renglones_Comprobantes.Documento = Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("           AND Comprobantes.Tipo = " & lcTipoComprobante & " AND Comprobantes.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			)")
            loComandoSeleccionar.AppendLine("LEFT JOIN  #tmpSaldoInicial ")
            loComandoSeleccionar.AppendLine("		ON  #tmpSaldoInicial.Cod_Aux = Renglones_Comprobantes.Cod_Aux ")
            loComandoSeleccionar.AppendLine("		AND #tmpSaldoInicial.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")
            loComandoSeleccionar.AppendLine("LEFT JOIN  #tmpMovimientos ")
            loComandoSeleccionar.AppendLine("		ON  #tmpMovimientos.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")

            loComandoSeleccionar.AppendLine("WHERE	    Renglones_Comprobantes.Fec_Ini  BETWEEN " & lcFechaDesde & " AND " & lcFechaHasta)
            loComandoSeleccionar.AppendLine("		AND	Renglones_Comprobantes.Cod_Cue	BETWEEN " & lcCuentaContableDesde & " AND " & lcCuentaContableHasta)
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Cen	BETWEEN " & lcCentroCostoDesde & " AND " & lcCentroCostoHasta)
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Gas	BETWEEN " & lcCuentaGastoDesde & " AND " & lcCuentaGastoHasta)
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Aux	BETWEEN " & lcAuxiliarDesde & " AND " & lcAuxiliarHasta)
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Cod_Mon	BETWEEN " & lcMonedaDesde & " AND " & lcMonedaHasta)
            'loComandoSeleccionar.AppendLine("ORDER BY Renglones_Comprobantes.Cod_Aux ")
            loComandoSeleccionar.AppendLine("")




            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT     Saldos.Posicion, ")
            loComandoSeleccionar.AppendLine("           Saldos.Saldo_Inicial, ")
            loComandoSeleccionar.AppendLine("           Saldos.Cod_Aux, ")
            loComandoSeleccionar.AppendLine("           Saldos.Debe, ")
            loComandoSeleccionar.AppendLine("           Saldos.Haber, ")
            loComandoSeleccionar.AppendLine("           Saldos.Saldo_Inicial + Saldos.Debe - Saldos.Haber + COALESCE(SUM(Finales.Debe - Finales.Haber), 0) as Saldo_Actual ")
            loComandoSeleccionar.AppendLine("INTO       #Saldos_Actuales ")
            loComandoSeleccionar.AppendLine("FROM       #Saldos AS Saldos ")
            loComandoSeleccionar.AppendLine("LEFT JOIN  #Saldos AS Finales ")
            loComandoSeleccionar.AppendLine("           ON Saldos.Cod_Aux = Finales.Cod_Aux ")
            loComandoSeleccionar.AppendLine("           AND Finales.Posicion < Saldos.Posicion")
            loComandoSeleccionar.AppendLine("GROUP BY   Saldos.Posicion,")
            loComandoSeleccionar.AppendLine("           Saldos.Saldo_Inicial, ")
            loComandoSeleccionar.AppendLine("           Saldos.Cod_Aux, ")
            loComandoSeleccionar.AppendLine("           Saldos.Debe, ")
            loComandoSeleccionar.AppendLine("           Saldos.Haber ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT     #Saldos.Documento                AS Documento,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Fec_Ini                  AS Fec_Ini,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Renglon                  AS Renglon,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Comentario               AS Comentario,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Debe                     AS Debe,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Haber                    AS Haber,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Cod_Aux                  AS Cod_Aux,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Cod_Cue                  AS Cod_Cue,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Nom_Cue                  AS Nom_Cue,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Posicion                 AS Posicion,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Saldo_Inicial            AS Saldo_Inicial, ")
            loComandoSeleccionar.AppendLine("           #Saldos.Cod_Niv_1                AS Cod_Niv_1,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Cod_Niv_2                AS Cod_Niv_2,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Cod_Niv_3                AS Cod_Niv_3,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Cod_Niv_4                AS Cod_Niv_4,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Cod_Niv_5                AS Cod_Niv_5,  ")
            loComandoSeleccionar.AppendLine("           #Saldos.Cod_Niv_6                AS Cod_Niv_6,  ")
            loComandoSeleccionar.AppendLine("           #Saldos_Actuales.Saldo_Actual    AS Saldo_Actual, ")
            loComandoSeleccionar.AppendLine("           COALESCE(Auxiliares.Nom_Aux, ")
            loComandoSeleccionar.AppendLine("            '[Sin Auxiliar Asignado]')      AS Nom_Aux ")
            loComandoSeleccionar.AppendLine("FROM       #Saldos  ")
            loComandoSeleccionar.AppendLine("LEFT JOIN  Auxiliares ")
            loComandoSeleccionar.AppendLine("           ON #Saldos.Cod_Aux = Auxiliares.Cod_Aux ")
            loComandoSeleccionar.AppendLine("LEFT JOIN  #Saldos_Actuales ")
            loComandoSeleccionar.AppendLine("           ON #Saldos_Actuales.Cod_Aux = #Saldos.Cod_Aux ")
            loComandoSeleccionar.AppendLine("           AND   #Saldos_Actuales.Posicion = #Saldos.Posicion")
            If Not llSoloMovimientos Then
                loComandoSeleccionar.AppendLine("UNION ALL ")
                loComandoSeleccionar.AppendLine("SELECT     NULL                             AS Documento, ")
                loComandoSeleccionar.AppendLine("           NULL                             AS Fec_Ini, ")
                loComandoSeleccionar.AppendLine("           NULL                             AS Renglon, ")
                loComandoSeleccionar.AppendLine("           '[Sin movimientos en el periodo]'AS Comentario, ")
                loComandoSeleccionar.AppendLine("           0                                AS Debe, ")
                loComandoSeleccionar.AppendLine("           0                                AS Haber, ")
                loComandoSeleccionar.AppendLine("           #tmpSaldoInicial.Cod_Aux         AS Cod_Aux, ")
                loComandoSeleccionar.AppendLine("           #tmpSaldoInicial.Cod_Cue         AS Cod_Cue, ")
                loComandoSeleccionar.AppendLine("           Cuentas_Contables.Nom_Cue        AS Nom_Cue, ")
                loComandoSeleccionar.AppendLine("           1                                AS Posicion, ")
                loComandoSeleccionar.AppendLine("           #tmpSaldoInicial.Inicial         AS Saldo_Inicial, ")
                loComandoSeleccionar.AppendLine("           #tmpMovimientos.Cod_Niv_1        AS Cod_Niv_1, ")
                loComandoSeleccionar.AppendLine("           #tmpMovimientos.Cod_Niv_2        AS Cod_Niv_2, ")
                loComandoSeleccionar.AppendLine("           #tmpMovimientos.Cod_Niv_3        AS Cod_Niv_3, ")
                loComandoSeleccionar.AppendLine("           #tmpMovimientos.Cod_Niv_4        AS Cod_Niv_4, ")
                loComandoSeleccionar.AppendLine("           #tmpMovimientos.Cod_Niv_5        AS Cod_Niv_5, ")
                loComandoSeleccionar.AppendLine("           #tmpMovimientos.Cod_Niv_6        AS Cod_Niv_6, ")
                loComandoSeleccionar.AppendLine("           #tmpSaldoInicial.Inicial         AS Saldo_Actual, ")
                loComandoSeleccionar.AppendLine("           Auxiliares.Nom_Aux               AS Nom_Aux ")
                loComandoSeleccionar.AppendLine("FROM        #tmpSaldoInicial")
                loComandoSeleccionar.AppendLine("LEFT JOIN  Auxiliares ")
                loComandoSeleccionar.AppendLine("           ON Auxiliares.Cod_Aux = #tmpSaldoInicial.Cod_Aux")
                loComandoSeleccionar.AppendLine("JOIN  #tmpMovimientos ")
                loComandoSeleccionar.AppendLine("           ON #tmpMovimientos.Cod_Cue = #tmpSaldoInicial.Cod_Cue")
                loComandoSeleccionar.AppendLine("JOIN  Cuentas_Contables ")
                loComandoSeleccionar.AppendLine("           ON Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Cue")
                loComandoSeleccionar.AppendLine("WHERE NOT EXISTS(   SELECT  * ")
                loComandoSeleccionar.AppendLine("                    FROM    #Saldos ")
                loComandoSeleccionar.AppendLine("                    WHERE   #Saldos.cod_cue = #tmpSaldoInicial.Cod_Cue ")
                loComandoSeleccionar.AppendLine("                        AND #Saldos.Cod_Aux = #tmpSaldoInicial.Cod_Aux)")
            End If
            loComandoSeleccionar.AppendLine("ORDER BY   Cod_Aux, Posicion")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")



            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes", 1800)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CCMV_rEstado_Cuenta_Auxiliares", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCCMV_rEstado_Cuenta_Auxiliares.ReportSource = loObjetoReporte

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

    Private Sub mConsultaOriginal()
        Return
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
        'loComandoSeleccionar.AppendLine("SET		@lcParametro1Desde = " & lcParametro0Desde) 'fecha
        'loComandoSeleccionar.AppendLine("SET		@lcParametro1Hasta = " & lcParametro0Hasta)	
        'loComandoSeleccionar.AppendLine("SET		@lcParametro2Desde = " & lcParametro1Desde)	'Cuenta Contable
        'loComandoSeleccionar.AppendLine("SET		@lcParametro2Hasta = " & lcParametro1Hasta)
        'loComandoSeleccionar.AppendLine("SET		@lcParametro3Desde = " & lcParametro2Desde)	'Centro de Costo
        'loComandoSeleccionar.AppendLine("SET		@lcParametro3Hasta = " & lcParametro2Hasta)
        'loComandoSeleccionar.AppendLine("SET		@lcParametro4Desde = " & lcParametro3Desde)	'Cunta de Gasto
        'loComandoSeleccionar.AppendLine("SET		@lcParametro4Hasta = " & lcParametro3Hasta)
        'loComandoSeleccionar.AppendLine("SET		@lcParametro5Desde = " & lcParametro4Desde)	'Auxiliares
        'loComandoSeleccionar.AppendLine("SET		@lcParametro5Hasta = " & lcParametro4Hasta)
        'loComandoSeleccionar.AppendLine("SET		@lcParametro6Desde = " & lcParametro5Desde)	'Moneda
        'loComandoSeleccionar.AppendLine("SET		@lcParametro6Hasta = " & lcParametro5Hasta)
        'loComandoSeleccionar.AppendLine("SET		@lcParametro7Desde = " & lcParametro6Desde)	'Tipo de Comprobante
        'loComandoSeleccionar.AppendLine("SET		@lcParametro8Desde = " & lcParametro7Desde)	'Con Movimientos
        'loComandoSeleccionar.AppendLine("SET		@lcParametro9Desde = " & lcParametro8Desde)	'Nivel
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("--*********************************************************")
        'loComandoSeleccionar.AppendLine("-- Prepara un listado de las cuentas contables a incluir  *")
        'loComandoSeleccionar.AppendLine("--*********************************************************")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10);")
        'loComandoSeleccionar.AppendLine("SET		@lnCero = 0;")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("DECLARE @lnTamaño1 INT; SET @lnTamaño1= " & lnTamaño1.ToString() & ";")
        'loComandoSeleccionar.AppendLine("DECLARE @lnTamaño2 INT; SET @lnTamaño2= " & lnTamaño2.ToString() & ";")
        'loComandoSeleccionar.AppendLine("DECLARE @lnTamaño3 INT; SET @lnTamaño3= " & lnTamaño3.ToString() & ";")
        'loComandoSeleccionar.AppendLine("DECLARE @lnTamaño4 INT; SET @lnTamaño4= " & lnTamaño4.ToString() & ";")
        'loComandoSeleccionar.AppendLine("DECLARE @lnTamaño5 INT; SET @lnTamaño5= " & lnTamaño5.ToString() & ";")
        'loComandoSeleccionar.AppendLine("DECLARE @lnTamaño6 INT; SET @lnTamaño6= " & lnTamaño6.ToString() & ";")
        'loComandoSeleccionar.AppendLine("DECLARE @lnTamaño7 INT; SET @lnTamaño7= " & lnTamaño7.ToString() & ";")
        'loComandoSeleccionar.AppendLine("DECLARE @lnTamañoMax INT; SET @lnTamañoMax = @lnTamaño7;")
        'loComandoSeleccionar.AppendLine("DECLARE @llUsarTasa BIT; SET @llUsarTasa = 0;")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("IF (@lcParametro9Desde = 1)   ")
        'loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño1")
        'loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde = 2)")
        'loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño2")
        'loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde = 3)")
        'loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño3 ")
        'loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde = 4)")
        'loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño4 ")
        'loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde = 5)")
        'loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño5;")
        'loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde = 6)")
        'loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño6;")
        'loComandoSeleccionar.AppendLine("ELSE IF (@lcParametro9Desde >= 7)")
        'loComandoSeleccionar.AppendLine("	SET @lnTamañoMax = @lnTamaño7;")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("DECLARE @lnTasaInicio DECIMAL(28, 10); SET @lnTasaInicio = 1; ")
        'loComandoSeleccionar.AppendLine("DECLARE @lnMovimientos BIT; ")
        'loComandoSeleccionar.AppendLine("SET @lnMovimientos = (CASE WHEN @lcParametro8Desde = 'Si' THEN 1 ELSE 0 END);")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("-- ********************************************************")
        'loComandoSeleccionar.AppendLine("-- Obtiene los saldos iniciales                           *")
        'loComandoSeleccionar.AppendLine("-- ********************************************************")
        'loComandoSeleccionar.AppendLine("SELECT		RC.Cod_Cue As Cod_Cue,")
        'loComandoSeleccionar.AppendLine("			SUM(RC.mon_deb - RC.mon_hab)	AS Saldo,")
        'loComandoSeleccionar.AppendLine("			RC.cod_cen, RC.cod_gas, RC.cod_aux, RC.cod_mon")
        'loComandoSeleccionar.AppendLine("INTO		#tmpSaldos")
        'loComandoSeleccionar.AppendLine("FROM		Renglones_Comprobantes AS RC")
        'loComandoSeleccionar.AppendLine("	INNER JOIN Comprobantes	AS C")
        'loComandoSeleccionar.AppendLine("		ON (	RC.Documento = C.Documento")
        'loComandoSeleccionar.AppendLine("				AND RC.Adicional = C.Adicional")
        'loComandoSeleccionar.AppendLine("				AND C.Tipo = @lcParametro7Desde ")
        'loComandoSeleccionar.AppendLine("				AND C.fec_ini< @lcParametro1Desde ")
        'loComandoSeleccionar.AppendLine("				AND C.Status <> 'Anulado'")
        'loComandoSeleccionar.AppendLine("				AND ABS(RC.mon_deb - RC.mon_hab)>0")
        'loComandoSeleccionar.AppendLine("			)")
        'loComandoSeleccionar.AppendLine("WHERE	RC.cod_cue BETWEEN  @lcParametro2Desde AND @lcParametro2Hasta")
        'loComandoSeleccionar.AppendLine("GROUP BY RC.Cod_Cue, RC.cod_cen, RC.cod_gas, RC.cod_aux, RC.cod_mon")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("SELECT	Cod_Cue, SUM(Saldo) As Saldo")
        'loComandoSeleccionar.AppendLine("INTO	#tmpSaldosProcesados ")
        'loComandoSeleccionar.AppendLine("FROM	#tmpSaldos")
        'loComandoSeleccionar.AppendLine("WHERE	cod_cen BETWEEN @lcParametro3Desde AND @lcParametro3Hasta")
        'loComandoSeleccionar.AppendLine("	AND cod_gas BETWEEN @lcParametro4Desde AND @lcParametro4Hasta")
        'loComandoSeleccionar.AppendLine("	AND cod_aux BETWEEN @lcParametro5Desde AND @lcParametro5Hasta")
        'loComandoSeleccionar.AppendLine("	AND cod_mon BETWEEN @lcParametro6Desde AND @lcParametro6Hasta")
        'loComandoSeleccionar.AppendLine("GROUP BY Cod_Cue")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("DROP TABLE #tmpSaldos")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("-- ********************************************************")
        'loComandoSeleccionar.AppendLine("-- Obtiene los movimientos actuales")
        'loComandoSeleccionar.AppendLine("-- ********************************************************")
        'loComandoSeleccionar.AppendLine("DECLARE @lnMaximaLongitud AS INT;")
        'loComandoSeleccionar.AppendLine("SET @lnMaximaLongitud = (SELECT MAX(LEN(SUBSTRING(Cuentas_Contables.Cod_Cue,1,@lnTamañoMax))) FROM Cuentas_Contables)")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("SELECT		Cuentas.Cod_Cue																					AS Cod_Cue,")
        'loComandoSeleccionar.AppendLine("			Cuentas.Nom_Cue																					AS Nom_Cue,")
        'loComandoSeleccionar.AppendLine("			Cuentas.Movimiento																				AS Movimiento,")
        'loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas.Cod_Cue, 1, CASE WHEN @lnTamaño1 <= @lnTamañoMax THEN @lnTamaño1 ELSE 0 END)	AS Nivel_1,")
        'loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas.Cod_Cue, 1, CASE WHEN @lnTamaño2 <= @lnTamañoMax THEN @lnTamaño2 ELSE 0 END)	AS Nivel_2,")
        'loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas.Cod_Cue, 1, CASE WHEN @lnTamaño3 <= @lnTamañoMax THEN @lnTamaño3 ELSE 0 END)	AS Nivel_3,")
        'loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas.Cod_Cue, 1, CASE WHEN @lnTamaño4 <= @lnTamañoMax THEN @lnTamaño4 ELSE 0 END)	AS Nivel_4,")
        'loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas.Cod_Cue, 1, CASE WHEN @lnTamaño5 <= @lnTamañoMax THEN @lnTamaño5 ELSE 0 END)	AS Nivel_5,")
        'loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas.Cod_Cue, 1, CASE WHEN @lnTamaño6 <= @lnTamañoMax THEN @lnTamaño6 ELSE 0 END)	AS Nivel_6,")
        'loComandoSeleccionar.AppendLine("			SUBSTRING(Cuentas.Cod_Cue, 1, CASE WHEN @lnTamaño7 <= @lnTamañoMax THEN @lnTamaño7 ELSE 0 END)	AS Nivel_7,")
        'loComandoSeleccionar.AppendLine("			CASE Cuentas.Cod_Cue")
        'loComandoSeleccionar.AppendLine("				WHEN  SUBSTRING(Cuentas.Cod_Cue, 1, @lnTamaño1) THEN 1 ")
        'loComandoSeleccionar.AppendLine("				WHEN  SUBSTRING(Cuentas.Cod_Cue, 1, @lnTamaño2) THEN 2 ")	 
        'loComandoSeleccionar.AppendLine("				WHEN  SUBSTRING(Cuentas.Cod_Cue, 1, @lnTamaño3) THEN 3 ")	 
        'loComandoSeleccionar.AppendLine("				WHEN  SUBSTRING(Cuentas.Cod_Cue, 1, @lnTamaño4) THEN 4 ")
        'loComandoSeleccionar.AppendLine("				WHEN  SUBSTRING(Cuentas.Cod_Cue, 1, @lnTamaño5) THEN 5 ")
        'loComandoSeleccionar.AppendLine("				WHEN  SUBSTRING(Cuentas.Cod_Cue, 1, @lnTamaño6) THEN 6 ")	 
        'loComandoSeleccionar.AppendLine("				WHEN  SUBSTRING(Cuentas.Cod_Cue, 1, @lnTamaño7) THEN 7 ")	 	 
        'loComandoSeleccionar.AppendLine("				ELSE 10")	 
        'loComandoSeleccionar.AppendLine("			END																								AS Nivel,")
        'loComandoSeleccionar.AppendLine("			@lnMaximaLongitud																				AS Maximo,")	 
        'loComandoSeleccionar.AppendLine("			MAX(ISNULL(Saldos.Saldo, @lnCero)) 																AS Saldo_Inicial,")
        'loComandoSeleccionar.AppendLine("			SUM(ISNULL(RC.mon_deb, @lnCero))																AS Debe,")
        'loComandoSeleccionar.AppendLine("			SUM(ISNULL(RC.mon_hab, @lnCero))																AS Haber,")
        'loComandoSeleccionar.AppendLine("			SUM(ISNULL(RC.mon_deb - RC.mon_hab, @lnCero))													AS Monto,")
        'loComandoSeleccionar.AppendLine("			MAX(ISNULL(Saldos.Saldo, @lnCero)) + SUM(ISNULL(RC.mon_deb - RC.mon_hab, @lnCero))				AS Saldo_Actual")
        'loComandoSeleccionar.AppendLine("INTO		#tmpListadoAgrupado")
        'loComandoSeleccionar.AppendLine("FROM		Renglones_Comprobantes AS RC")
        'loComandoSeleccionar.AppendLine("	INNER JOIN Comprobantes	AS C")
        'loComandoSeleccionar.AppendLine("		ON (	RC.Documento = C.Documento")
        'loComandoSeleccionar.AppendLine("				AND RC.Adicional = C.Adicional")
        'loComandoSeleccionar.AppendLine("				AND C.Tipo = @lcParametro7Desde ")
        'loComandoSeleccionar.AppendLine("				AND RC.cod_cue BETWEEN @lcParametro2Desde AND @lcParametro2Hasta")
        'loComandoSeleccionar.AppendLine("				AND RC.cod_cen BETWEEN @lcParametro3Desde AND @lcParametro3Hasta")
        'loComandoSeleccionar.AppendLine("				AND RC.cod_gas BETWEEN @lcParametro4Desde AND @lcParametro4Hasta")
        'loComandoSeleccionar.AppendLine("				AND RC.cod_aux BETWEEN @lcParametro5Desde AND @lcParametro5Hasta")
        'loComandoSeleccionar.AppendLine("				AND RC.cod_mon BETWEEN @lcParametro6Desde AND @lcParametro6Hasta")
        'loComandoSeleccionar.AppendLine("				AND RC.fec_ini BETWEEN @lcParametro1Desde AND @lcParametro1Hasta")
        'loComandoSeleccionar.AppendLine("				AND C.Status <> 'Anulado'")
        'loComandoSeleccionar.AppendLine("			)")
        'loComandoSeleccionar.AppendLine("	RIGHT JOIN (")
        'loComandoSeleccionar.AppendLine("				SELECT	Nombres.Cod_Cue,")
        'loComandoSeleccionar.AppendLine("						Nombres.Nom_Cue,")
        'loComandoSeleccionar.AppendLine("						Nombres.Movimiento")
        'loComandoSeleccionar.AppendLine("				FROM	Cuentas_Contables AS Nombres")
        'loComandoSeleccionar.AppendLine("					JOIN (	")
        'loComandoSeleccionar.AppendLine("							SELECT	SUBSTRING(Cuentas_Contables.Cod_Cue,1,@lnTamañoMax) AS Cod_Cue")
        'loComandoSeleccionar.AppendLine("							FROM	Cuentas_Contables ")
        'loComandoSeleccionar.AppendLine("						)AS Codigos ON Codigos.Cod_Cue = Nombres.Cod_Cue")
        'loComandoSeleccionar.AppendLine("				WHERE Nombres.Cod_Cue BETWEEN @lcParametro2Desde AND @lcParametro2Hasta")
        'loComandoSeleccionar.AppendLine("				GROUP BY Nombres.Cod_Cue,Nombres.Nom_Cue,Nombres.Movimiento")
        'loComandoSeleccionar.AppendLine("			) AS Cuentas")
        'loComandoSeleccionar.AppendLine("			ON	Cuentas.Cod_Cue = SUBSTRING(RC.Cod_Cue,1,@lnTamañoMax) ")
        'loComandoSeleccionar.AppendLine("	LEFT JOIN	#tmpSaldosProcesados AS Saldos")
        'loComandoSeleccionar.AppendLine("			ON	Cuentas.Cod_Cue = SUBSTRING(Saldos.Cod_Cue,1,@lnTamañoMax) ")
        'loComandoSeleccionar.AppendLine("GROUP BY  Cuentas.Cod_Cue, Cuentas.Nom_Cue, Cuentas.Movimiento")
        'If llSoloMovimientos Then 
        '	loComandoSeleccionar.AppendLine("HAVING	ABS(SUM(ISNULL(Saldos.Saldo, @lnCero))) ")
        '	loComandoSeleccionar.AppendLine("	  + ABS(SUM(ISNULL(RC.Mon_Deb, @lnCero))) ")
        '	loComandoSeleccionar.AppendLine("	  + ABS(SUM(ISNULL(RC.Mon_Hab, @lnCero))) > 0")
        '	loComandoSeleccionar.AppendLine("	OR Cuentas.Movimiento = 0 ")
        'End If
        'loComandoSeleccionar.AppendLine("ORDER BY Cuentas.Cod_Cue")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("")
        'loComandoSeleccionar.AppendLine("DROP TABLE #tmpSaldosProcesados")
        'loComandoSeleccionar.AppendLine("")

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' GSC: 04/09/12: Codigo inicial.									                        '
'-------------------------------------------------------------------------------------------' 
' RJG: 17/12/13: Ajuste en filtros el cálculo de saldos iniciales.                          '
'-------------------------------------------------------------------------------------------' 
' RJG: 19/12/13: Se incluyeron los auxiliares con saldo inicial que no tengan movimientos   '
'                en el periodo.                                                             '
'-------------------------------------------------------------------------------------------' 
