'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rRConciliacion_Bancaria"
'-------------------------------------------------------------------------------------------'


Partial Class CGS_rRConciliacion_Bancaria

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''' OBTIENE EL SALDO INICIAL CONCILIADO ANTES DE LA FECHA INDICADA EN LOS PARAMETROS''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            loComandoSeleccionar.AppendLine("DECLARE @lcCuenta      	VARCHAR(10)")
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaInicio		DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaFin		DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE @ldSaldoEstado		DECIMAL(28, 10)")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero			DECIMAL(28,10)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET @lcCuenta      		= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("SET @ldFechaInicio			= " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("SET @ldFechaFin			= " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("SET @ldSaldoEstado			= " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("SET @lnCero				= 0")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Cuentas_Bancarias.Cod_Cue		AS Cod_Cue,")
            loComandoSeleccionar.AppendLine("        Cuentas_Bancarias.Num_Cue		AS Num_Cue,")
            loComandoSeleccionar.AppendLine("        Bancos.Nom_Ban					AS Nom_Ban")
            loComandoSeleccionar.AppendLine("INTO #tmpCuenta")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Bancarias")
            loComandoSeleccionar.AppendLine("    JOIN Bancos ON Cuentas_Bancarias.Cod_Ban = Bancos.Cod_Ban")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Bancarias.Cod_Cue = @lcCuenta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Saldo Inicial Conciliado")
            loComandoSeleccionar.AppendLine("SELECT		Movimientos_Cuentas.Cod_Cue										AS Cod_Cue, ")
            loComandoSeleccionar.AppendLine("       	SUM(Movimientos_Cuentas.Mon_Deb) 								AS Mon_Deb_Ini,")
            loComandoSeleccionar.AppendLine("       	SUM(Movimientos_Cuentas.Mon_Hab) 								AS Mon_Hab_Ini,")
            loComandoSeleccionar.AppendLine("       	SUM(Movimientos_Cuentas.Mon_Deb - Movimientos_Cuentas.Mon_Hab)	AS Mon_Sal_Ini")
            loComandoSeleccionar.AppendLine("INTO		#tmpSaldoInicialConciliado")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cuentas.Doc_Cil  = 1")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Fec_Ini < @ldFechaInicio")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Cod_Cue = @lcCuenta")
            loComandoSeleccionar.AppendLine("GROUP BY	Movimientos_Cuentas.Cod_Cue")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Movimientos Conciliados")
            loComandoSeleccionar.AppendLine("SELECT		Movimientos_Cuentas.Cod_Cue									AS Cod_Cue, ")
            loComandoSeleccionar.AppendLine("			SUM(Movimientos_Cuentas.Mon_Deb)		 					AS Cargos,")
            loComandoSeleccionar.AppendLine("			SUM(Movimientos_Cuentas.Mon_Hab) 							AS Abonos,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpSaldoInicialConciliado.Mon_Sal_Ini, @lnCero)		AS Sal_Ini_Con,")
            loComandoSeleccionar.AppendLine("			(	  SUM(Movimientos_Cuentas.Mon_Deb) ")
            loComandoSeleccionar.AppendLine("				- SUM(Movimientos_Cuentas.Mon_Hab) ")
            loComandoSeleccionar.AppendLine("				+ ISNULL(#tmpSaldoInicialConciliado.Mon_Sal_Ini, @lnCero)")
            loComandoSeleccionar.AppendLine("			)															AS Sal_Fin_Con,")
            loComandoSeleccionar.AppendLine("			@lnCero														AS Sal_Fin_No_Con")
            loComandoSeleccionar.AppendLine("INTO		#tmpConciliado")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpSaldoInicialConciliado ON (#tmpSaldoInicialConciliado.Cod_Cue = Movimientos_Cuentas.Cod_Cue)")
            loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cuentas.Doc_Cil = 1 ")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Cod_Cue = @lcCuenta")
            loComandoSeleccionar.AppendLine("GROUP BY	Movimientos_Cuentas.Cod_Cue,")
            loComandoSeleccionar.AppendLine("			#tmpSaldoInicialConciliado.Mon_Deb_Ini, #tmpSaldoInicialConciliado.Mon_Hab_Ini, #tmpSaldoInicialConciliado.Mon_Sal_Ini")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpSaldoInicialConciliado")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Movimientos No Conciliados")
            loComandoSeleccionar.AppendLine("SELECT		Movimientos_Cuentas.Cod_Cue			AS Cod_Cue,")
            loComandoSeleccionar.AppendLine("			@lnCero								AS Sal_Ini_Con,		")
            loComandoSeleccionar.AppendLine("			SUM(Movimientos_Cuentas.Mon_Deb)	AS CargosNC,")
            loComandoSeleccionar.AppendLine("			SUM(Movimientos_Cuentas.Mon_Hab)	AS AbonosNC,")
            loComandoSeleccionar.AppendLine("			@lnCero								AS Sal_Fin_Con,")
            loComandoSeleccionar.AppendLine("			SUM(Movimientos_Cuentas.Mon_Deb) - ")
            loComandoSeleccionar.AppendLine("				SUM(Movimientos_Cuentas.Mon_Hab)AS Sal_Fin_No_Con")
            loComandoSeleccionar.AppendLine("INTO		#tmpNoConciliado")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cuentas.Doc_Cil = 0")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Cod_Cue = @lcCuenta")
            loComandoSeleccionar.AppendLine("GROUP BY	Movimientos_Cuentas.Cod_Cue")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Movimientos en Tránsito")
            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tipo,")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro")
            loComandoSeleccionar.AppendLine(" INTO #tmpTransito")
            loComandoSeleccionar.AppendLine(" FROM  Movimientos_Cuentas ")
            loComandoSeleccionar.AppendLine(" JOIN Ordenes_Pagos ON Movimientos_Cuentas.Doc_Ori = Ordenes_Pagos.Documento")
            loComandoSeleccionar.AppendLine("   AND Movimientos_Cuentas.Tip_Ori = 'Ordenes_Pagos'")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("   AND Movimientos_Cuentas.Doc_Ori = Ordenes_Pagos.Documento")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Cue = @lcCuenta")
            loComandoSeleccionar.AppendLine("           AND (Movimientos_Cuentas.Fec_Cil NOT BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("           OR Movimientos_Cuentas.Doc_Cil     =   '0') ")

            loComandoSeleccionar.AppendLine("UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Tipo,")
            loComandoSeleccionar.AppendLine("           Proveedores.Nom_Pro")
            loComandoSeleccionar.AppendLine(" FROM  Movimientos_Cuentas ")
            loComandoSeleccionar.AppendLine(" JOIN Pagos ON Movimientos_Cuentas.Doc_Ori = Pagos.Documento")
            loComandoSeleccionar.AppendLine("   AND Movimientos_Cuentas.Tip_Ori = 'Pagos'")
            loComandoSeleccionar.AppendLine(" JOIN Proveedores ON Proveedores.Cod_Pro = Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("   AND Movimientos_Cuentas.Doc_Ori = Pagos.Documento")
            loComandoSeleccionar.AppendLine(" WHERE     Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Cue = @lcCuenta")
            loComandoSeleccionar.AppendLine("           AND (Movimientos_Cuentas.Fec_Cil NOT BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("           OR Movimientos_Cuentas.Doc_Cil     =   '0') ")

            loComandoSeleccionar.AppendLine("UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT    Movimientos_Cuentas.Documento, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Referencia, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           CASE WHEN Movimientos_Cuentas.Cod_Tip = 'Cre-TP' ")
            loComandoSeleccionar.AppendLine("           THEN 'Transferencia' ELSE Movimientos_Cuentas.Tipo END AS Tipo,")
            loComandoSeleccionar.AppendLine("           Movimientos_Cuentas.Comentario AS Nom_Pro")
            loComandoSeleccionar.AppendLine(" FROM  Movimientos_Cuentas ")
            loComandoSeleccionar.AppendLine(" JOIN  Cuentas_Bancarias ON Movimientos_Cuentas.Cod_Cue     =   Cuentas_Bancarias.Cod_Cue")
            loComandoSeleccionar.AppendLine(" WHERE     (Movimientos_Cuentas.Cod_Tip = 'Cre-TP' OR Movimientos_Cuentas.Cod_Tip = 'Cre-NC')")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("           AND Movimientos_Cuentas.Cod_Cue = @lcCuenta")
            loComandoSeleccionar.AppendLine("           AND (Movimientos_Cuentas.Fec_Cil NOT BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("           OR Movimientos_Cuentas.Doc_Cil     =   '0') ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		#tmpCuenta.Cod_Cue									    AS Cod_Cue,")
            loComandoSeleccionar.AppendLine("           #tmpCuenta.Num_Cue									    AS Num_Cue,")
            loComandoSeleccionar.AppendLine("			#tmpCuenta.Nom_Ban									    AS Nom_Ban,")
            loComandoSeleccionar.AppendLine("			#tmpConciliado.Sal_Ini_Con								AS Sal_Ini_Con,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpConciliado.Cargos,@lnCero) 					AS Cargos,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpConciliado.Abonos,@lnCero) 					AS Abonos,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpConciliado.Sal_Fin_Con,@lnCero)				AS Sal_Fin_Con,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpNoConciliado.CargosNC,@lnCero) 				AS CargosNC,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpNoConciliado.AbonosNC,@lnCero) 				AS AbonosNC,")
            loComandoSeleccionar.AppendLine("			ISNULL((#tmpNoConciliado.CargosNC - ")
            loComandoSeleccionar.AppendLine("				#tmpNoConciliado.AbonosNC),@lnCero)					AS Sal_Fin_No_Con,")
            loComandoSeleccionar.AppendLine("			@ldSaldoEstado											AS Sal_Edo_Cta,	")
            loComandoSeleccionar.AppendLine("			(ISNULL(#tmpConciliado.Sal_Fin_Con, @lnCero) - ")
            loComandoSeleccionar.AppendLine("				@ldSaldoEstado)										AS Diferencia,	")
            loComandoSeleccionar.AppendLine("			ISNULL((#tmpConciliado.Sal_Fin_Con + ")
            loComandoSeleccionar.AppendLine("               #tmpNoConciliado.Sal_Fin_No_Con),@lnCero)           AS Saldo_Libro,")
            loComandoSeleccionar.AppendLine("			ISNULL((#tmpConciliado.Abonos + ")
            loComandoSeleccionar.AppendLine("               #tmpNoConciliado.AbonosNC),@lnCero)                 AS Saldo_Abonos,")
            loComandoSeleccionar.AppendLine("			ISNULL((#tmpConciliado.Cargos + ")
            loComandoSeleccionar.AppendLine("               #tmpNoConciliado.CargosNC),@lnCero)                 AS Saldo_Cargos,")
            loComandoSeleccionar.AppendLine("           #tmpTransito.Documento, ")
            loComandoSeleccionar.AppendLine("           #tmpTransito.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           CONCAT(RTRIM(#tmpTransito.Tipo), ' ',RTRIM(#tmpTransito.Referencia)) AS Referencia,")
            loComandoSeleccionar.AppendLine("           #tmpTransito.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           #tmpTransito.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           #tmpTransito.Nom_Pro")
            loComandoSeleccionar.AppendLine("FROM		#tmpConciliado")
            loComandoSeleccionar.AppendLine("	FULL JOIN #tmpNoConciliado ON #tmpConciliado.Cod_Cue =  #tmpNoConciliado.Cod_Cue")
            loComandoSeleccionar.AppendLine("   FULL JOIN #tmpTransito ON #tmpConciliado.Cod_Cue =  #tmpTransito.Cod_Cue")
            loComandoSeleccionar.AppendLine("   FULL JOIN #tmpCuenta ON #tmpConciliado.Cod_Cue = #tmpCuenta.Cod_Cue")
            loComandoSeleccionar.AppendLine("ORDER BY #tmpConciliado.Cod_Cue")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpConciliado")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpNoConciliado")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpTransito ")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCuenta")
            loComandoSeleccionar.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rRConciliacion_Bancaria", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrMCuentas_Conciliados.ReportSource = loObjetoReporte

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
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 21/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GCR: 27/03/09: Ajustes al diseño
'-------------------------------------------------------------------------------------------'
' CMS: 04/07/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS: 04/07/09: Se agregaron los siguientes filtros: Tipos de movimientos, Revisión, Sucursal
'-------------------------------------------------------------------------------------------'
' RJG: 23/11/11: Ajuste en el los filtros de fechas del SELECT.								'
'-------------------------------------------------------------------------------------------'
