﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rEGanancias_Perdidas"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rEGanancias_Perdidas

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim lcFechaDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcFechaHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcCuentaContableDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcCuentaContableHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            'Dim lcAuxiliarDesde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            'Dim lcAuxiliarHasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim llSoloMovimientos As Boolean = CStr(cusAplicacion.goReportes.paParametrosIniciales(2)).Trim().ToUpper().Equals("SI")
            Dim lnNivelMax As Integer = CInt(cusAplicacion.goReportes.paParametrosIniciales(3))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

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

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10) = 0")
            loComandoSeleccionar.AppendLine("DECLARE @lcFechaDesde DATETIME = " & lcFechaDesde)
            loComandoSeleccionar.AppendLine("DECLARE @lcFechaHasta DATETIME = " & lcFechaHasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCuentaDesde VARCHAR(30) = " & lcCuentaContableDesde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCuentaHasta VARCHAR(30) = " & lcCuentaContableHasta)
            'loComandoSeleccionar.AppendLine("DECLARE @lcAuxDesde VARCHAR(30) = " & lcAuxiliarDesde)
            'loComandoSeleccionar.AppendLine("DECLARE @lcAuxHasta VARCHAR(30) = " & lcAuxiliarHasta)
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
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini < @lcFechaDesde)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Saldo,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>= @lcFechaDesde AND Renglones_Comprobantes.Fec_Ini<= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Debe,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>= @lcFechaDesde AND Renglones_Comprobantes.Fec_Ini<= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Haber,")
            loComandoSeleccionar.AppendLine("		SUM(CASE	WHEN (Renglones_Comprobantes.Fec_Ini>= @lcFechaDesde AND Renglones_Comprobantes.Fec_Ini<= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("					THEN ISNULL(Renglones_Comprobantes.Mon_Deb - Renglones_Comprobantes.Mon_Hab, @lnCero)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("			END")
            loComandoSeleccionar.AppendLine("		) AS Monto")
            loComandoSeleccionar.AppendLine("INTO	#tmpMovimientos")
            loComandoSeleccionar.AppendLine("FROM	Cuentas_Contables AS CC")
            loComandoSeleccionar.AppendLine("	LEFT OUTER JOIN Renglones_Comprobantes ")
            loComandoSeleccionar.AppendLine("		INNER JOIN Comprobantes")
            loComandoSeleccionar.AppendLine("			ON (Renglones_Comprobantes.Adicional = Comprobantes.Adicional ")
            loComandoSeleccionar.AppendLine("				AND Renglones_Comprobantes.Documento = Comprobantes.Documento")
            loComandoSeleccionar.AppendLine("				AND Comprobantes.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			)")
            loComandoSeleccionar.AppendLine("		ON CC.Cod_Cue = Renglones_Comprobantes.Cod_Cue ")
            loComandoSeleccionar.AppendLine("			AND (Renglones_Comprobantes.Fec_Ini <= @lcFechaHasta)")
            loComandoSeleccionar.AppendLine("WHERE	CC.Movimiento=1")
            loComandoSeleccionar.AppendLine("   AND CC.Categoria NOT IN ('Activos', 'Pasivos', 'Capital')")
            loComandoSeleccionar.AppendLine("	AND	CC.Cod_Cue						BETWEEN @lcCuentaDesde	AND	@lcCuentaHasta")
            'loComandoSeleccionar.AppendLine("	AND Renglones_Comprobantes.Cod_Aux	BETWEEN @lcAuxDesde	AND	@lcAuxHasta")
            loComandoSeleccionar.AppendLine("GROUP BY CC.Cod_Cue")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Cuentas_Contables.Cod_Cue							AS Cod_Cue,")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Nom_Cue							AS Nom_Cue,")
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
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Debe)							AS Debe,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Haber)							AS Haber,")
            loComandoSeleccionar.AppendLine("			SUM(#tmpMovimientos.Monto + #tmpMovimientos.Saldo)	AS Saldo_Actual,")
            loComandoSeleccionar.AppendLine("           @lcFechaDesde                                       AS Desde,")
            loComandoSeleccionar.AppendLine("           @lcFechaHasta                                       AS Hasta")
            loComandoSeleccionar.AppendLine("FROM		#tmpMovimientos")
            loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Contables")
            loComandoSeleccionar.AppendLine("		ON	Cuentas_Contables.Cod_Cue = #tmpMovimientos.Cod_Cue")
            loComandoSeleccionar.AppendLine("GROUP BY	Cuentas_Contables.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Nom_Cue, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Contables.Movimiento,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_1, ")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_2, ")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_3,")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_4, ")
            loComandoSeleccionar.AppendLine("			#tmpMovimientos.Cod_Niv_5 ")
            If Not (llSoloMovimientos) Then
                loComandoSeleccionar.AppendLine("HAVING	ABS(SUM(#tmpMovimientos.Monto + #tmpMovimientos.Saldo)) > 0")
            End If
            loComandoSeleccionar.AppendLine("ORDER BY	Cod_Cue")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpMovimientos")
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

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rEGanancias_Perdidas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvTRF_rEGanancias_Perdidas.ReportSource = loObjetoReporte

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
' RJG: 26/10/11: Codigo inicial, a partir de Libro Diario.									'
'-------------------------------------------------------------------------------------------' 
' RJG: 27/10/11: Corrección en cálculo de saldos inicial y final. Corrección en filtro.	'
'-------------------------------------------------------------------------------------------' 
' RJG: 16/11/11: Agregado filtro de la estructura superior cuando en el detalle no hay		'
'				 movimientos (y el usuario indicó el filtro "Solo Moviminetos = SI".		'
'-------------------------------------------------------------------------------------------' 
' RJG: 06/12/11: Se agregó la igualdad de campo Adicional en las uniones entre Comprobantes	'
'				 y sus renglones.															'
'-------------------------------------------------------------------------------------------' 
' RJG: 09/02/12: Cambiado el filtro de comprobantes "='Pendiente'" por "<>'Anulado'".		'
'-------------------------------------------------------------------------------------------' 
