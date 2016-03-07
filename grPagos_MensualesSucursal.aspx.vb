'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grPagos_MensualesSucursal"
'-------------------------------------------------------------------------------------------'
Partial Class grPagos_MensualesSucursal
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT 1 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes INTO #tempFECHAS")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 2 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 3 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 4 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 5 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 6 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 7 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 8 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 9 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 10 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 11 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 12 AS Mes," & lcParametro0Desde & " AS Año, 0 As Pago_Mes")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine("       Sucursales.Cod_Suc,")
            loComandoSeleccionar.AppendLine(" 		Sucursales.Nom_Suc,")
            loComandoSeleccionar.AppendLine(" 		#tempFECHAS.Mes,")
            loComandoSeleccionar.AppendLine(" 		#tempFECHAS.Año,")
            loComandoSeleccionar.AppendLine(" 		#tempFECHAS.Pago_Mes")
            loComandoSeleccionar.AppendLine(" INTO	#tempRESULTBASIC")
            loComandoSeleccionar.AppendLine(" FROM	Sucursales")
            loComandoSeleccionar.AppendLine(" CROSS JOIN #tempFECHAS")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Sucursales.Cod_Suc,")
            loComandoSeleccionar.AppendLine(" 		Sucursales.Nom_Suc,")
            loComandoSeleccionar.AppendLine("       DATEPART(MONTH, Pagos.fec_ini)AS Mes,")
            loComandoSeleccionar.AppendLine("       DATEPART(YEAR, Pagos.fec_ini) AS Año,")
            loComandoSeleccionar.AppendLine("       SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) AS Pago_Mes")
            loComandoSeleccionar.AppendLine(" INTO #tempRESULT")
            loComandoSeleccionar.AppendLine(" FROM Pagos")
            loComandoSeleccionar.AppendLine(" JOIN Sucursales ON  Sucursales.Cod_Suc = Pagos.Cod_Suc")
            loComandoSeleccionar.AppendLine(" JOIN Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine(" WHERE DATEPART(YEAR, Pagos.Fec_ini) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Rev BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Suc BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Sucursales.Cod_Suc, Sucursales.Nom_Suc, DATEPART(YEAR, Pagos.fec_ini), DATEPART(MONTH, Pagos.fec_ini)")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Sucursales.Cod_Suc,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=1) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_Ene,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=2) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_Feb,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=3) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_Mar,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=4) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_Abr,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=5) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_May,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=6) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_Jun,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=7) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_Jul,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=8) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_Ago,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=9) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_Sep,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=10) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_Oct,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=11) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_Nov,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Pagos.fec_ini)=12) THEN SUM(ISNULL(Detalles_Pagos.Mon_Net,0)) ELSE 0 END AS Monto_Dic")
            loComandoSeleccionar.AppendLine(" INTO #tempRESULT2")
            loComandoSeleccionar.AppendLine(" FROM Pagos")
            loComandoSeleccionar.AppendLine(" JOIN Sucursales ON  Sucursales.Cod_Suc = Pagos.Cod_Suc")
            loComandoSeleccionar.AppendLine(" JOIN Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine(" WHERE DATEPART(YEAR, Pagos.Fec_ini) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Rev BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       AND Pagos.Cod_Suc BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Sucursales.Cod_Suc, Sucursales.Nom_Suc, DATEPART(YEAR, Pagos.fec_ini), DATEPART(MONTH, Pagos.fec_ini)")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		#tempRESULTBASIC.Cod_Suc,")
            loComandoSeleccionar.AppendLine(" 		#tempRESULTBASIC.Nom_Suc,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULTBASIC.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine(" 		END AS Str_Mes,")
            loComandoSeleccionar.AppendLine("       #tempRESULTBASIC.Mes,")
            loComandoSeleccionar.AppendLine("       #tempRESULTBASIC.Año,")
            loComandoSeleccionar.AppendLine("       #tempRESULTBASIC.Pago_Mes")
            loComandoSeleccionar.AppendLine(" INTO #tempRESULTFINAL")
            loComandoSeleccionar.AppendLine(" FROM #tempRESULTBASIC")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		#tempRESULT.Cod_Suc,")
            loComandoSeleccionar.AppendLine(" 		#tempRESULT.Nom_Suc,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine("           WHEN #tempRESULT.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine(" 		END AS Str_Mes,")
            loComandoSeleccionar.AppendLine("       #tempRESULT.Mes,")
            loComandoSeleccionar.AppendLine("       #tempRESULT.Año,     ")
            loComandoSeleccionar.AppendLine("       #tempRESULT.Pago_Mes")
            loComandoSeleccionar.AppendLine(" FROM #tempRESULT")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		#tempRESULTFINAL.Cod_Suc,")
            loComandoSeleccionar.AppendLine(" 		#tempRESULTFINAL.Nom_Suc,")
            loComandoSeleccionar.AppendLine(" 		#tempRESULTFINAL.Str_Mes,")
            loComandoSeleccionar.AppendLine("       #tempRESULTFINAL.Mes,")
            loComandoSeleccionar.AppendLine("       #tempRESULTFINAL.Año,")
            loComandoSeleccionar.AppendLine("       SUM(#tempRESULTFINAL.Pago_Mes) AS Pago_Mes,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_Ene,0) AS Monto_Ene,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_Feb,0) AS Monto_Feb,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_Mar,0) AS Monto_Mar,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_Abr,0) AS Monto_Abr,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_May,0) AS Monto_May,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_Jun,0) AS Monto_Jun,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_Jul,0) AS Monto_Jul,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_Ago,0) AS Monto_Ago,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_Sep,0) AS Monto_Sep,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_Oct,0) AS Monto_Oct,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_Nov,0) AS Monto_Nov,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tempRESULT2.Monto_Dic,0) AS Monto_Dic,")
            loComandoSeleccionar.AppendLine(" 		SUM(ISNULL(#tempRESULT2.Monto_Ene,0) + ISNULL(#tempRESULT2.Monto_Feb,0) + ISNULL(#tempRESULT2.Monto_Mar,0) + ")
            loComandoSeleccionar.AppendLine(" 			ISNULL(#tempRESULT2.Monto_Abr,0) + ISNULL(#tempRESULT2.Monto_May,0) + ISNULL(#tempRESULT2.Monto_Jun,0) + ")
            loComandoSeleccionar.AppendLine(" 			ISNULL(#tempRESULT2.Monto_Jul,0) + ISNULL(#tempRESULT2.Monto_Ago,0) + ISNULL(#tempRESULT2.Monto_Sep,0) + ")
            loComandoSeleccionar.AppendLine(" 			ISNULL(#tempRESULT2.Monto_Oct,0) + ISNULL(#tempRESULT2.Monto_Nov,0) + ISNULL(#tempRESULT2.Monto_Dic,0)) AS Tot_Pago")
            loComandoSeleccionar.AppendLine(" FROM #tempRESULTFINAL")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN #tempRESULT2 ON #tempRESULTFINAL.Cod_Suc = #tempRESULT2.Cod_Suc")
            loComandoSeleccionar.AppendLine(" GROUP BY #tempRESULTFINAL.Cod_Suc, #tempRESULTFINAL.Nom_Suc, #tempRESULTFINAL.Str_Mes, #tempRESULTFINAL.Mes, #tempRESULTFINAL.Año,")
            loComandoSeleccionar.AppendLine("       #tempRESULT2.Monto_Ene, #tempRESULT2.Monto_Feb, #tempRESULT2.Monto_Mar, #tempRESULT2.Monto_Abr,")
            loComandoSeleccionar.AppendLine("       #tempRESULT2.Monto_May, #tempRESULT2.Monto_Jun, #tempRESULT2.Monto_Jul, #tempRESULT2.Monto_Ago,")
            loComandoSeleccionar.AppendLine("       #tempRESULT2.Monto_Sep, #tempRESULT2.Monto_Oct, #tempRESULT2.Monto_Nov, #tempRESULT2.Monto_Dic")
            loComandoSeleccionar.AppendLine(" ORDER BY #tempRESULTFINAL.Cod_Suc, #tempRESULTFINAL.Año, #tempRESULTFINAL.Mes")

'me.mEscribirConsulta(loComandoSeleccionar.ToString)
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grPagos_MensualesSucursal", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrPagos_MensualesSucursal.ReportSource = loObjetoReporte

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
' CMS: 29/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
