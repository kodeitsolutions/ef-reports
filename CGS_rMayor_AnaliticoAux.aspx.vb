'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rMayor_AnaliticoAux"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rMayor_AnaliticoAux
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
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecIni			AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecFin			AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_CodCue_Desde	AS VARCHAR(30)")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_CodCue_Hasta	AS VARCHAR(30)")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_Auxiliar_Desde	AS VARCHAR(30)")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_Auxiliar_Hasta	AS VARCHAR(30)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero AS DECIMAL(28, 10)")
            loComandoSeleccionar.AppendLine("SET @lnCero = 0;	")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET	@sp_FecIni = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("SET	@sp_FecFin = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("SET    @sp_CodCue_Desde = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("SET    @sp_CodCue_Hasta = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("SET    @sp_Auxiliar_Desde = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("SET    @sp_Auxiliar_Hasta = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Saldo Inicial")
            loComandoSeleccionar.AppendLine("SELECT	Renglones_Comprobantes.Cod_Cue		AS Cod_Cuenta,")
            loComandoSeleccionar.AppendLine("       Renglones_Comprobantes.Cod_Aux		AS Cod_Aux,")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab)	AS Saldo_Inicial,")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Comprobantes.Mon_Deb)	AS Saldo_Debe,")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Comprobantes.Mon_Hab)	AS Saldo_Haber")
            loComandoSeleccionar.AppendLine("INTO	#tmpSaldos_Iniciales")
            loComandoSeleccionar.AppendLine("FROM	Renglones_Comprobantes ")
            loComandoSeleccionar.AppendLine("	INNER JOIN Comprobantes ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("		AND Comprobantes.Tipo = 'Diario' AND Comprobantes.Status <> 'Anulado')")
            loComandoSeleccionar.AppendLine("   JOIN Cuentas_Contables ON Renglones_Comprobantes.Cod_Cue = Cuentas_Contables.Cod_Cue")
            loComandoSeleccionar.AppendLine("WHERE	Renglones_Comprobantes.Fec_Ini < @sp_FecIni")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Cod_Cue BETWEEN @sp_CodCue_Desde AND @sp_CodCue_Hasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Cod_Aux BETWEEN @sp_Auxiliar_Desde AND @sp_Auxiliar_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Contables.Auxiliar = 1")
            loComandoSeleccionar.AppendLine("GROUP BY Renglones_Comprobantes.Cod_Cue, Renglones_Comprobantes.Cod_Aux")
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
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Contables ON Renglones_Comprobantes.Cod_Cue = Cuentas_Contables.Cod_Cue")
            loComandoSeleccionar.AppendLine("WHERE Renglones_Comprobantes.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("   AND Renglones_Comprobantes.Cod_Cue BETWEEN @sp_CodCue_Desde AND @sp_CodCue_Hasta ")
            loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Cod_Aux BETWEEN @sp_Auxiliar_Desde AND @sp_Auxiliar_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Contables.Auxiliar = 1")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE	#tmpMovimientos")
            loComandoSeleccionar.AppendLine("SET	Saldo_Inicial = COALESCE(Mov.Saldo_Inicial,0),")
            loComandoSeleccionar.AppendLine("		Saldo = CASE WHEN Mov.Categoria IN ('Capital','Ingresos', 'Pasivos')")
            loComandoSeleccionar.AppendLine("					THEN -Mov.Debe + Mov.Haber")
            loComandoSeleccionar.AppendLine("					ELSE Mov.Debe - Mov.Haber")
            loComandoSeleccionar.AppendLine("				END,")
            loComandoSeleccionar.AppendLine("		Orden = Mov.Orden")
            loComandoSeleccionar.AppendLine("FROM	(SELECT ROW_NUMBER () OVER (PARTITION BY #tmpMovimientos.Cod_Cuenta, #tmpMovimientos.Auxiliar ORDER BY #tmpMovimientos.Fecha, #tmpMovimientos.Auxiliar ASC) AS Orden,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Documento, #tmpMovimientos.Renglon, #tmpMovimientos.Cod_Cuenta, #tmpMovimientos.Auxiliar,")
            loComandoSeleccionar.AppendLine("           Saldo_Ini.Saldo_Inicial,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Categoria,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Debe,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Haber")
            loComandoSeleccionar.AppendLine("		FROM #tmpMovimientos")
            loComandoSeleccionar.AppendLine("			LEFT JOIN (SELECT Cod_Cuenta, Cod_Aux, SUM(Saldo_Debe) AS Saldo_Debe, SUM(Saldo_Haber) AS Saldo_Haber, Saldo_Inicial FROM #tmpSaldos_Iniciales GROUP BY Cod_Cuenta, Cod_Aux, Saldo_Inicial) AS Saldo_Ini")
            loComandoSeleccionar.AppendLine("				ON Saldo_Ini.Cod_Cuenta = #tmpMovimientos.Cod_Cuenta AND Saldo_Ini.Cod_Aux = #tmpMovimientos.Auxiliar")
            loComandoSeleccionar.AppendLine("		) AS Mov")
            loComandoSeleccionar.AppendLine("WHERE Mov.Documento = #tmpMovimientos.Documento")
            loComandoSeleccionar.AppendLine("	AND Mov.Renglon = #tmpMovimientos.Renglon")
            loComandoSeleccionar.AppendLine("	AND Mov.Cod_Cuenta = #tmpMovimientos.Cod_Cuenta")
            loComandoSeleccionar.AppendLine("   AND Mov.Auxiliar = #tmpMovimientos.Auxiliar")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	A.Orden, A.Documento, A.Fecha_Comprobante, A.Renglon, A.Cod_Cuenta, A.Nom_Cuenta, A.Auxiliar,")
            loComandoSeleccionar.AppendLine("		Auxiliares.Nom_Aux, A.Saldo_Inicial, A.Categoria, A.Fecha, A.Debe, A.Haber,")
            loComandoSeleccionar.AppendLine("       CASE WHEN A.Categoria IN ('Capital','Ingresos','Pasivos')")
            loComandoSeleccionar.AppendLine("       	 THEN A.Saldo_Inicial - (SUM(B.Saldo))")
            loComandoSeleccionar.AppendLine("       	 ELSE  A.Saldo_Inicial + SUM(B.Saldo)")
            loComandoSeleccionar.AppendLine("       END								AS Sal_Doc,")
            loComandoSeleccionar.AppendLine("       @sp_FecIni AS Desde, @sp_FecFin	AS Hasta, A.Comentario")
            loComandoSeleccionar.AppendLine("FROM	#tmpMovimientos AS A")
            loComandoSeleccionar.AppendLine("	JOIN #tmpMovimientos AS B ON B.Cod_Cuenta = A.Cod_Cuenta")
            loComandoSeleccionar.AppendLine("       AND A.Auxiliar = B.Auxiliar")
            loComandoSeleccionar.AppendLine("		AND B.Orden <= A.Orden")
            loComandoSeleccionar.AppendLine("   JOIN Auxiliares ON Auxiliares.Cod_Aux = A.Auxiliar")
            loComandoSeleccionar.AppendLine("GROUP BY A.Orden, A.Documento, A.Fecha_Comprobante, A.Renglon, A.Cod_Cuenta, A.Nom_Cuenta, Auxiliares.Nom_Aux,")
            loComandoSeleccionar.AppendLine("		A.Auxiliar, A.Saldo_Inicial, A.Categoria, A.Fecha, A.Debe, A.Haber, A.Saldo_Inicial, A.Comentario")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Cuenta ASC, Auxiliar ASC, A.Fecha ASC")
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rMayor_AnaliticoAux", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvCGS_rMayor_AnaliticoAux.ReportSource = loObjetoReporte

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
