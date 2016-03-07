'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRConciliacion_Bancaria"
'-------------------------------------------------------------------------------------------'


Partial Class rRConciliacion_Bancaria

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(Strings.Format(CInt(cusAplicacion.goReportes.paParametrosIniciales(1)), "0000"))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(Strings.Format(CInt(cusAplicacion.goReportes.paParametrosIniciales(2)), "00"))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio)
							
			If lcParametro1Desde = "'0000'"  Then 
				lcParametro1Desde =	goServicios.mObtenerCampoFormatoSQL(Strings.Format(Date.Today().Year(), "0000"))
			ElseIf lcParametro1Desde < "'1900'"  Then 
				lcParametro1Desde =	"'1900'"
			End If
			
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
			
            Dim loComandoSeleccionar As New StringBuilder()          
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''' OBTIENE EL SALDO INICIAL CONCILIADO ANTES DE LA FECHA INDICADA EN LOS PARAMETROS''''
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            loComandoSeleccionar.AppendLine("DECLARE @lcCuentaInicio	VARCHAR(10)")
            loComandoSeleccionar.AppendLine("DECLARE @lcCuentaFin		VARCHAR(10)")
            loComandoSeleccionar.AppendLine("DECLARE @lcAño				VARCHAR(4)")
            loComandoSeleccionar.AppendLine("DECLARE @lcMes				VARCHAR(2)")
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaInicio		DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaFin		DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero			DECIMAL(28,10)")
            loComandoSeleccionar.AppendLine("DECLARE @ldSaldoEstado		DECIMAL(28, 10)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET @lcCuentaInicio		= " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("SET @lcCuentaFin			= " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("SET @lcAño					= " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("SET @lcMes					= " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("SET @ldFechaInicio			= CAST(@lcAño + @lcMes + '01' AS DATETIME)")
            loComandoSeleccionar.AppendLine("SET @ldFechaFin			= DATEADD(MS, -3, DATEADD(MONTH, 1, @ldFechaInicio)) ")
            loComandoSeleccionar.AppendLine("SET @lnCero				= 0")
            loComandoSeleccionar.AppendLine("SET @ldSaldoEstado			= " & lcParametro3Desde)
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
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin")
            loComandoSeleccionar.AppendLine("GROUP BY	Movimientos_Cuentas.Cod_Cue")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Movimientos Conciliados")
            loComandoSeleccionar.AppendLine("SELECT		Movimientos_Cuentas.Cod_Cue									AS Cod_Cue, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Bancarias.Num_Cue									AS Num_Cue,")
            loComandoSeleccionar.AppendLine("			Bancos.Nom_Ban												AS Nom_Ban,")
            loComandoSeleccionar.AppendLine("			Movimientos_Cuentas.Cod_Mon									AS Cod_Mon,")
            loComandoSeleccionar.AppendLine("			SUM(Movimientos_Cuentas.Mon_Deb)		 					AS Cargos,")
            loComandoSeleccionar.AppendLine("			SUM(Movimientos_Cuentas.Mon_Hab) 							AS Abonos,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpSaldoInicialConciliado.Mon_Sal_Ini, @lnCero)		AS Sal_Ini_Con,")
            loComandoSeleccionar.AppendLine("			(	  SUM(Movimientos_Cuentas.Mon_Deb) ")
            loComandoSeleccionar.AppendLine("				- SUM(Movimientos_Cuentas.Mon_Hab) ")
            loComandoSeleccionar.AppendLine("				+ ISNULL(#tmpSaldoInicialConciliado.Mon_Sal_Ini, @lnCero)")
            loComandoSeleccionar.AppendLine("			)															AS Sal_Fin_Con,")
            loComandoSeleccionar.AppendLine("			'Conciliado'												AS Tipo, ")
            loComandoSeleccionar.AppendLine("			@lnCero														AS Sal_Fin_No_Con")
            loComandoSeleccionar.AppendLine("INTO		#tmpConciliado")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Bancarias ON (Movimientos_Cuentas.Cod_Cue = Cuentas_Bancarias.Cod_Cue)")
            loComandoSeleccionar.AppendLine("	JOIN	Bancos ON (Cuentas_Bancarias.Cod_Ban = Bancos.Cod_Ban)")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpSaldoInicialConciliado ON (#tmpSaldoInicialConciliado.Cod_Cue = Movimientos_Cuentas.Cod_Cue)")
            loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cuentas.Doc_Cil = 1 ")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin")
            loComandoSeleccionar.AppendLine("GROUP BY	Movimientos_Cuentas.Cod_Cue, Cuentas_Bancarias.Num_Cue, Bancos.Nom_Ban,Movimientos_Cuentas.Cod_Mon,")
            loComandoSeleccionar.AppendLine("			#tmpSaldoInicialConciliado.Mon_Deb_Ini, #tmpSaldoInicialConciliado.Mon_Hab_Ini, #tmpSaldoInicialConciliado.Mon_Sal_Ini")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpSaldoInicialConciliado")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Movimientos No Conciliados")
            loComandoSeleccionar.AppendLine("SELECT		Movimientos_Cuentas.Cod_Cue			AS Cod_Cue,")
            loComandoSeleccionar.AppendLine("			Cuentas_Bancarias.Num_Cue			AS Num_Cue,")
            loComandoSeleccionar.AppendLine("			Bancos.Nom_Ban						AS Nom_Ban,")
            loComandoSeleccionar.AppendLine("			Movimientos_Cuentas.Cod_Mon			AS Cod_Mon,")
            loComandoSeleccionar.AppendLine("			@lnCero								AS Sal_Ini_Con,		")
            loComandoSeleccionar.AppendLine("			SUM(Movimientos_Cuentas.Mon_Deb)	AS CargosNC,")
            loComandoSeleccionar.AppendLine("			SUM(Movimientos_Cuentas.Mon_Hab)	AS AbonosNC,")
            loComandoSeleccionar.AppendLine("			@lnCero								AS Sal_Fin_Con,")
            loComandoSeleccionar.AppendLine("			'No_Conciliado'						AS Tipo,")
            loComandoSeleccionar.AppendLine("			SUM(Movimientos_Cuentas.Mon_Deb) - ")
            loComandoSeleccionar.AppendLine("				SUM(Movimientos_Cuentas.Mon_Hab)AS Sal_Fin_No_Con")
            loComandoSeleccionar.AppendLine("INTO		#tmpNoConciliado")
            loComandoSeleccionar.AppendLine("FROM		Movimientos_Cuentas")
            loComandoSeleccionar.AppendLine(" 	JOIN	Cuentas_Bancarias ON (Movimientos_Cuentas.Cod_Cue = Cuentas_Bancarias.Cod_Cue)")
            loComandoSeleccionar.AppendLine(" 	JOIN	Bancos ON (Cuentas_Bancarias.Cod_Ban = Bancos.Cod_Ban)")
            loComandoSeleccionar.AppendLine("WHERE		Movimientos_Cuentas.Doc_Cil = 0")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Fec_Ini BETWEEN @ldFechaInicio AND @ldFechaFin")
            loComandoSeleccionar.AppendLine("		AND Movimientos_Cuentas.Cod_Cue BETWEEN @lcCuentaInicio AND @lcCuentaFin")
            loComandoSeleccionar.AppendLine("GROUP BY	Movimientos_Cuentas.Cod_Cue, Cuentas_Bancarias.Num_Cue, Bancos.Nom_Ban, Movimientos_Cuentas.Cod_Mon")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		#tmpConciliado.Cod_Cue									AS Cod_Cue,")
            loComandoSeleccionar.AppendLine("			#tmpConciliado.Num_Cue									AS Num_Cue,")
            loComandoSeleccionar.AppendLine("			#tmpConciliado.Nom_Ban									AS Nom_Ban,")
            loComandoSeleccionar.AppendLine("			#tmpConciliado.Cod_Mon									AS Cod_Mon,")
            loComandoSeleccionar.AppendLine("			#tmpConciliado.Tipo										AS Tipo,	")
            loComandoSeleccionar.AppendLine("			#tmpConciliado.Sal_Ini_Con								AS Sal_Ini_Con,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpConciliado.Cargos,0) 						AS Cargos,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpConciliado.Abonos,0) 						AS Abonos,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpConciliado.Sal_Fin_Con,0)					AS Sal_Fin_Con,")
            loComandoSeleccionar.AppendLine("			#tmpNoConciliado.Cod_Cue 								AS Cod_CueNC,")
            loComandoSeleccionar.AppendLine("			#tmpNoConciliado.Num_Cue 								AS Num_CueNC,")
            loComandoSeleccionar.AppendLine("			#tmpNoConciliado.Nom_Ban 								AS Nom_BanNC,")
            loComandoSeleccionar.AppendLine("			#tmpNoConciliado.Cod_Mon 								AS Cod_MonNC,")
            loComandoSeleccionar.AppendLine("			#tmpNoConciliado.Tipo									AS TipoNC,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpNoConciliado.CargosNC,0) 					AS CargosNC,")
            loComandoSeleccionar.AppendLine("			ISNULL(#tmpNoConciliado.AbonosNC,0) 					AS AbonosNC,")
            loComandoSeleccionar.AppendLine("			ISNULL((#tmpNoConciliado.CargosNC - ")
            loComandoSeleccionar.AppendLine("				#tmpNoConciliado.AbonosNC),0)						AS Sal_Fin_No_Con,")
            loComandoSeleccionar.AppendLine("			@ldSaldoEstado											AS Sal_Edo_Cta,	")
            loComandoSeleccionar.AppendLine("			(ISNULL(#tmpConciliado.Sal_Fin_Con, @lnCero) - ")
            loComandoSeleccionar.AppendLine("				@ldSaldoEstado)										AS Diferencia	")
            loComandoSeleccionar.AppendLine("FROM		#tmpConciliado")
            loComandoSeleccionar.AppendLine("	FULL JOIN #tmpNoConciliado ON #tmpConciliado.Cod_Cue =  #tmpNoConciliado.Cod_Cue")
            loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpConciliado")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpNoConciliado")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
			
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
            

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRConciliacion_Bancaria", laDatosReporte)

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
