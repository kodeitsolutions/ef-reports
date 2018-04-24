'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rMayor_Analitico"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rMayor_Analitico
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecIni			AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecFin			AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE	@sp_CodCue_Desde	AS VARCHAR(30) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE	@sp_CodCue_Hasta	AS VARCHAR(30) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE	@sp_Documento_Desde	AS VARCHAR(10) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE	@sp_Documento_Hasta	AS VARCHAR(10) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE	@sp_Auxiliar_Desde	AS VARCHAR(30) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE	@sp_Auxiliar_Hasta	AS VARCHAR(30) = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero AS DECIMAL(28, 10) = 0")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Saldo Inicial")
            loComandoSeleccionar.AppendLine("SELECT	Renglones_Comprobantes.Cod_Cue		AS Cod_Cuenta,")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab)	AS Saldo_Inicial,")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Comprobantes.Mon_Deb)	AS Saldo_Debe,")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Comprobantes.Mon_Hab)	AS Saldo_Haber")
            loComandoSeleccionar.AppendLine("INTO	#tmpSaldos_Iniciales")
            loComandoSeleccionar.AppendLine("FROM	Renglones_Comprobantes ")
            loComandoSeleccionar.AppendLine("	INNER JOIN Comprobantes ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("		AND Comprobantes.Tipo = 'Diario' AND Comprobantes.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("		AND Comprobantes.Documento BETWEEN @sp_Documento_Desde AND @sp_Documento_Hasta)")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Contables ON Renglones_Comprobantes.Cod_Cue = Cuentas_Contables.Cod_Cue")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Contables.Categoria IN ('Activos', 'Pasivos', 'Capital')")
            loComandoSeleccionar.AppendLine("WHERE	Renglones_Comprobantes.Fec_Ini < @sp_FecIni")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Cod_Cue BETWEEN @sp_CodCue_Desde AND @sp_CodCue_Hasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Documento BETWEEN @sp_Documento_Desde AND @sp_Documento_Hasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Cod_Aux BETWEEN @sp_Auxiliar_Desde AND @sp_Auxiliar_Hasta")
            loComandoSeleccionar.AppendLine("GROUP BY Renglones_Comprobantes.Cod_Cue")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	0									AS Orden,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Documento	AS Documento,")
            loComandoSeleccionar.AppendLine("       Renglones_Comprobantes.Fec_Ini      AS Fecha_Comprobante,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Renglon		AS Renglon,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Cue		AS Cod_Cuenta,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Nom_Cue		AS Nom_Cuenta,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Cod_Aux		AS Auxiliar,")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Categoria			AS Categoria,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Fec_Ori		AS Fecha,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Mon_Deb		AS Debe,")
            loComandoSeleccionar.AppendLine("		Renglones_Comprobantes.Mon_Hab		AS Haber,")
            loComandoSeleccionar.AppendLine("		@lnCero								AS Saldo,")
            loComandoSeleccionar.AppendLine("		@lnCero								AS Saldo_Doc,")
            loComandoSeleccionar.AppendLine("		@lnCero								AS Saldo_Inicial,")
            loComandoSeleccionar.AppendLine("       Renglones_Comprobantes.Comentario   AS Comentario")
            loComandoSeleccionar.AppendLine("INTO #tmpMovimientos")
            loComandoSeleccionar.AppendLine("FROM Renglones_Comprobantes")
            loComandoSeleccionar.AppendLine("	JOIN Comprobantes ON Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("		AND Comprobantes.Tipo = 'Diario' AND Comprobantes.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Contables ON Renglones_Comprobantes.Cod_Cue = Cuentas_Contables.Cod_Cue")
            loComandoSeleccionar.AppendLine("WHERE Renglones_Comprobantes.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Cod_Cue BETWEEN @sp_CodCue_Desde AND @sp_CodCue_Hasta ")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Documento BETWEEN @sp_Documento_Desde AND @sp_Documento_Hasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Cod_Aux BETWEEN @sp_Auxiliar_Desde AND @sp_Auxiliar_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE	#tmpMovimientos")
            loComandoSeleccionar.AppendLine("SET		Orden = Mov.Orden,")
            loComandoSeleccionar.AppendLine("		Saldo = CASE WHEN Mov.Categoria IN ('Capital','Ingresos', 'Pasivos')")
            loComandoSeleccionar.AppendLine("					THEN -Mov.Debe + Mov.Haber")
            loComandoSeleccionar.AppendLine("					ELSE Mov.Debe - Mov.Haber")
            loComandoSeleccionar.AppendLine("				END,")
            loComandoSeleccionar.AppendLine("		Saldo_Inicial = COALESCE(Mov.Saldo_Inicial,@lnCero)")
            loComandoSeleccionar.AppendLine("FROM	(SELECT ROW_NUMBER () OVER (PARTITION BY #tmpMovimientos.Cod_Cuenta ORDER BY #tmpMovimientos.Fecha ASC) AS Orden,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Documento, #tmpMovimientos.Renglon, #tmpMovimientos.Cod_Cuenta,")
            loComandoSeleccionar.AppendLine("           Saldo_Ini.Saldo_Inicial,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Categoria,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Debe,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Haber")
            loComandoSeleccionar.AppendLine("		FROM #tmpMovimientos")
            loComandoSeleccionar.AppendLine("			LEFT JOIN (SELECT Cod_Cuenta, SUM(Saldo_Debe) AS Saldo_Debe, SUM(Saldo_Haber) AS Saldo_Haber, Saldo_Inicial FROM #tmpSaldos_Iniciales GROUP BY Cod_Cuenta, Saldo_Inicial) AS Saldo_Ini")
            loComandoSeleccionar.AppendLine("				ON Saldo_Ini.Cod_Cuenta = #tmpMovimientos.Cod_Cuenta")
            loComandoSeleccionar.AppendLine("		) AS Mov")
            loComandoSeleccionar.AppendLine("WHERE Mov.Documento = #tmpMovimientos.Documento")
            loComandoSeleccionar.AppendLine("	AND Mov.Renglon = #tmpMovimientos.Renglon")
            loComandoSeleccionar.AppendLine("	AND Mov.Cod_Cuenta = #tmpMovimientos.Cod_Cuenta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	A.Orden, A.Documento, A.Fecha_Comprobante, A.Renglon, A.Cod_Cuenta, A.Nom_Cuenta, A.Auxiliar,")
            loComandoSeleccionar.AppendLine("		A.Saldo_Inicial, A.Categoria, A.Fecha, A.Debe, A.Haber,")
            loComandoSeleccionar.AppendLine("       CASE WHEN A.Categoria IN ('Capital','Ingresos','Pasivos')")
            loComandoSeleccionar.AppendLine("       	 THEN A.Saldo_Inicial - (SUM(B.Saldo))")
            loComandoSeleccionar.AppendLine("       	 ELSE  A.Saldo_Inicial + SUM(B.Saldo)")
            loComandoSeleccionar.AppendLine("       END								AS Sal_Doc,")
            loComandoSeleccionar.AppendLine("       @sp_FecIni AS Desde, @sp_FecFin	AS Hasta, A.Comentario," & lcParametro4Desde & " AS Agrupar")
            loComandoSeleccionar.AppendLine("FROM	#tmpMovimientos AS A")
            loComandoSeleccionar.AppendLine("	JOIN #tmpMovimientos AS B")
            loComandoSeleccionar.AppendLine("		ON B.Cod_Cuenta = A.Cod_Cuenta")
            loComandoSeleccionar.AppendLine("		AND B.Orden <= A.Orden")
            loComandoSeleccionar.AppendLine("GROUP BY A.Orden, A.Documento, A.Fecha_Comprobante, A.Renglon, A.Cod_Cuenta, A.Nom_Cuenta, ")
            loComandoSeleccionar.AppendLine("		A.Auxiliar, A.Saldo_Inicial, A.Categoria, A.Fecha, A.Debe, A.Haber, A.Saldo_Inicial, A.Comentario")
            loComandoSeleccionar.AppendLine("ORDER BY FECHA ASC, Renglon ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpSaldos_Iniciales")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rMayor_Analitico", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rMayor_Analitico.ReportSource = loObjetoReporte

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
' DLC: 02/09/2010: Programacion inicial (Replica del reporte rLEstadoCuenta_HistoricoVentas)'
'                   - Cambio de la consulta a procedimiento almacenado.						'
'-------------------------------------------------------------------------------------------'
' DLC: 15/09/2010: Ajuste en la forma de obtener los detalles de Pagos, asi como también,	'
'                ajustar en el RPT, la forma de mostrar los detalles de Pagos.				'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Reprogramación del Reporte y su respectivo Store Procedure					'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Ajuste de la vista de Diseño.												'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Se elimino el filtro Detalle												'
'-------------------------------------------------------------------------------------------'
' RJG: 05/12/11: Eliminado el SP: ahora la consulta se hace desde un Query en línea para	'
'				 corregir cálculo de saldo y optimizar.										'
'-------------------------------------------------------------------------------------------'
