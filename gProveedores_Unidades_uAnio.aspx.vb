'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "gProveedores_Unidades_uAnio"
'-------------------------------------------------------------------------------------------'
Partial Class gProveedores_Unidades_uAnio
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            '--------------------------------------------------------------------------------------
            'Para el cálculo de la linea de tendencia de cada grupo se sigue el siguiente
            'procedimiento de forma general:
            '
            'La recta representa : y = ( a * x ) + b
            '
            'donde:
            '   n = número de registros
            '   Ex = SUMATORIA ( xi ) / n
            '   Ey = SUMATOIRA ( yi ) / n
            '
            '   a = SUMATORIA( (xi - Ex) * (yi - Ey) ) / SUMATORIA( (xi - Ex) * (xi - Ex) )
            '   b = Ey - ( a * Ex )
            '
            'Ejemplo:
            '   x       y
            '   10      100
            '   21      200
            '   30      300
            '
            ' n = 3
            ' Ex = (10 + 21 + 30) / 3 = 61/3 = 20.33
            ' Ey = (100 + 200 + 300 ) / 3 = 600/3 = 200
            '
            ' SUMATORIA( (xi - Ex) * (yi - Ey) ) = ((10 - 20.33) * (100 - 200)) 
            '                                    + ((21 - 20.33) * (200 - 200)) 
            '                                    + ((30 - 20.33) * (300 - 200)) = 2000
            ' SUMATORIA( (xi - Ex) * (xi - Ex) ) = ((10 - 20.33) * (10 - 20.33)) 
            '                                    + ((21 - 20.33) * (21 - 20.33)) 
            '                                    + ((30 - 20.33) * (30 - 20.33)) = 200.6667
            ' a = 2000/200.6667 = 9.96677575
            ' b = 200 - ( 9.96677575 * 20.33) = -2.624551
            '
            ' entonces:
            '   x       y       tendencia
            '   10      100      97.0432065
            '   21      200     206.67774
            '   30      300     296.378722
            '--------------------------------------------------------------------------------------

            loComandoSeleccionar.AppendLine(" DECLARE @sum_Ey real")
            loComandoSeleccionar.AppendLine(" DECLARE @sum_Ex real")
            loComandoSeleccionar.AppendLine(" DECLARE @sum_xEx_yEy real")
            loComandoSeleccionar.AppendLine(" DECLARE @sum_xEx_xEx real")
            loComandoSeleccionar.AppendLine(" DECLARE @valor_a real")
            loComandoSeleccionar.AppendLine(" DECLARE @valor_b real")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT 1 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=1 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades INTO #temp_unicom")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 2 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=2 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 3 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=3 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 4 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=4 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 5 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=5 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 6 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=6 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 7 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=7 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 8 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=8 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 9 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=9 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 10 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=10 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 11 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=11 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 12 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=12 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		DATEPART(MONTH,Compras.Fec_Ini) AS Mes,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(YEAR,Compras.Fec_Ini) AS Anio,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Compras.Can_Art1 AS Unidades")
            loComandoSeleccionar.AppendLine(" FROM Proveedores")
            loComandoSeleccionar.AppendLine(" JOIN Compras ON Proveedores.Cod_Pro = Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Compras ON Compras.Documento = Renglones_Compras.Documento")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 		DATEDIFF(MONTH,Compras.Fec_Ini,GETDATE()) < 12")
            loComandoSeleccionar.AppendLine("       AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 	    #temp_unicom.Mes,")
            loComandoSeleccionar.AppendLine(" 	    #temp_unicom.Anio,")
            loComandoSeleccionar.AppendLine(" 	    SUM(#temp_unicom.Unidades) AS Unidades,")
            loComandoSeleccionar.AppendLine(" 	    ROW_NUMBER() OVER(ORDER BY #temp_unicom.Anio ASC, #temp_unicom.Mes ASC) AS Row")
            loComandoSeleccionar.AppendLine(" INTO #temp_totcom")
            loComandoSeleccionar.AppendLine(" FROM #temp_unicom")
            loComandoSeleccionar.AppendLine(" GROUP BY Mes, Anio")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SET @sum_Ey = (SELECT SUM(#temp_totcom.Unidades) FROM #temp_totcom)/12")
            loComandoSeleccionar.AppendLine(" SET @sum_Ex = 6.5")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		((#temp_totcom.Row-@sum_Ex)*(#temp_totcom.Unidades-@sum_Ey)) AS xEx_yEy,")
            loComandoSeleccionar.AppendLine(" 		((#temp_totcom.Row-@sum_Ex)*(#temp_totcom.Row-@sum_Ex)) AS xEx_xEx")
            loComandoSeleccionar.AppendLine(" INTO #temp_tencom")
            loComandoSeleccionar.AppendLine(" FROM #temp_totcom")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SET @sum_xEx_yEy = (SELECT SUM(#temp_tencom.xEx_yEy) FROM #temp_tencom)")
            loComandoSeleccionar.AppendLine(" SET @sum_xEx_xEx = (SELECT SUM(#temp_tencom.xEx_xEx) FROM #temp_tencom)")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SET @valor_a = (@sum_xEx_yEy / @sum_xEx_xEx)")
            loComandoSeleccionar.AppendLine(" SET @valor_b = (@sum_Ey-(@valor_a*@sum_Ex))")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 	    #temp_totcom.Mes,")
            loComandoSeleccionar.AppendLine(" 	    #temp_totcom.Anio,")
            loComandoSeleccionar.AppendLine(" 	    #temp_totcom.Unidades,")
            loComandoSeleccionar.AppendLine(" 	    ((@valor_a*#temp_totcom.Row)+@valor_b) AS Tendencia")
            loComandoSeleccionar.AppendLine(" INTO #temp_ttencom")
            loComandoSeleccionar.AppendLine(" FROM #temp_totcom")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT 1 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=1 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades INTO #temp_unidev")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 2 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=2 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 3 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=3 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 4 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=4 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 5 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=5 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 6 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=6 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 7 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=7 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 8 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=8 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 9 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=9 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 10 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=10 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 11 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=11 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 12 AS Mes,CASE WHEN DATEPART(MONTH,GETDATE())>=12 THEN DATEPART(YEAR,GETDATE()) ELSE DATEPART(YEAR,GETDATE())-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		DATEPART(MONTH,Devoluciones_Proveedores.Fec_Ini) AS Mes,")
            loComandoSeleccionar.AppendLine(" 		DATEPART(YEAR,Devoluciones_Proveedores.Fec_Ini) AS Anio,")
            loComandoSeleccionar.AppendLine(" 		Renglones_DProveedores.Can_Art1 AS Unidades")
            loComandoSeleccionar.AppendLine(" FROM Proveedores")
            loComandoSeleccionar.AppendLine(" JOIN Devoluciones_Proveedores ON Proveedores.Cod_Pro = Devoluciones_Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_DProveedores ON Devoluciones_Proveedores.Documento = Renglones_DProveedores.Documento")
            loComandoSeleccionar.AppendLine(" WHERE ")
            loComandoSeleccionar.AppendLine(" 		DATEDIFF(MONTH,Devoluciones_Proveedores.Fec_Ini,GETDATE()) < 12")
            loComandoSeleccionar.AppendLine("       AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 	    #temp_unidev.Mes,")
            loComandoSeleccionar.AppendLine(" 	    #temp_unidev.Anio,")
            loComandoSeleccionar.AppendLine(" 	    SUM(#temp_unidev.Unidades) AS Unidades,")
            loComandoSeleccionar.AppendLine(" 	    ROW_NUMBER() OVER(ORDER BY #temp_unidev.Anio ASC, #temp_unidev.Mes ASC) AS Row")
            loComandoSeleccionar.AppendLine(" INTO #temp_totdev")
            loComandoSeleccionar.AppendLine(" FROM #temp_unidev")
            loComandoSeleccionar.AppendLine(" GROUP BY Mes, Anio")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SET @sum_Ey = (SELECT SUM(#temp_totdev.Unidades) FROM #temp_totdev)/12")
            loComandoSeleccionar.AppendLine(" SET @sum_Ex = 6.5")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		((#temp_totdev.Row-@sum_Ex)*(#temp_totdev.Unidades-@sum_Ey)) AS xEx_yEy,")
            loComandoSeleccionar.AppendLine(" 		((#temp_totdev.Row-@sum_Ex)*(#temp_totdev.Row-@sum_Ex)) AS xEx_xEx")
            loComandoSeleccionar.AppendLine(" INTO #temp_tendev")
            loComandoSeleccionar.AppendLine(" FROM #temp_totdev")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SET @sum_xEx_yEy = (SELECT SUM(#temp_tendev.xEx_yEy) FROM #temp_tendev)")
            loComandoSeleccionar.AppendLine(" SET @sum_xEx_xEx = (SELECT SUM(#temp_tendev.xEx_xEx) FROM #temp_tendev)")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SET @valor_a = (@sum_xEx_yEy / @sum_xEx_xEx)")
            loComandoSeleccionar.AppendLine(" SET @valor_b = (@sum_Ey-(@valor_a*@sum_Ex))")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 	    #temp_totdev.Mes,")
            loComandoSeleccionar.AppendLine(" 	    #temp_totdev.Anio,")
            loComandoSeleccionar.AppendLine(" 	    #temp_totdev.Unidades,")
            loComandoSeleccionar.AppendLine(" 	    ((@valor_a*#temp_totdev.Row)+@valor_b) AS Tendencia")
            loComandoSeleccionar.AppendLine(" INTO #temp_ttendev")
            loComandoSeleccionar.AppendLine(" FROM #temp_totdev")
            loComandoSeleccionar.AppendLine(" ")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 	    Proveedores.Cod_Pro,")
            loComandoSeleccionar.AppendLine(" 	    Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine(" 	    #temp_ttencom.Mes AS Num_Mes,")
            loComandoSeleccionar.AppendLine(" 	    CASE")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine(" 		    WHEN #temp_ttencom.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine(" 	    END AS Str_Mes,")
            loComandoSeleccionar.AppendLine(" 	    #temp_ttencom.Anio AS Anio,")
            loComandoSeleccionar.AppendLine(" 	    #temp_ttencom.Unidades AS Compras,")
            loComandoSeleccionar.AppendLine(" 	    #temp_ttencom.Tendencia AS TendenciaCompras,")
            loComandoSeleccionar.AppendLine(" 	    #temp_ttendev.Unidades AS Devoluciones,")
            loComandoSeleccionar.AppendLine(" 	    #temp_ttendev.Tendencia AS TendenciaDevoluciones")
            loComandoSeleccionar.AppendLine(" FROM #temp_ttencom,#temp_ttendev, Proveedores")
            loComandoSeleccionar.AppendLine(" WHERE	")
            loComandoSeleccionar.AppendLine(" 	    #temp_ttencom.Mes = #temp_ttendev.Mes")
            loComandoSeleccionar.AppendLine("       AND " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("gProveedores_Unidades_uAnio", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgProveedores_Unidades_uAnio.ReportSource = loObjetoReporte

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
' Douglas Cortez 29/04/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
