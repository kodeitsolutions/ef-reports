'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grTransacciones_32Dias"
'-------------------------------------------------------------------------------------------'
Partial Class grTransacciones_32Dias
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            '--------------------------------------------------------------------------------------
            'Para el cálculo de la linea de tendencia de cada grupo se sigue el siguiente
            'procedimiento de forma general:
            '
            'La recta representa: y = ( a * x ) + b
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
            
            Dim ldFechaInicio as Date = DateTime.Parse(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim ldFechaFin as Date = DateTime.Parse(cusAplicacion.goReportes.paParametrosFinales(0))
			Dim LcAux As String = lcParametro0Hasta


            loComandoSeleccionar.AppendLine(" DECLARE @Sum_Ey Real")
            loComandoSeleccionar.AppendLine(" DECLARE @Sum_Ex Real")
            loComandoSeleccionar.AppendLine(" DECLARE @Sum_fac_xEx_yEy Real")
            loComandoSeleccionar.AppendLine(" DECLARE @Sum_fac_xEx_xEx Real")
            loComandoSeleccionar.AppendLine(" DECLARE @Valor_A Real")
            loComandoSeleccionar.AppendLine(" DECLARE @Valor_B Real")

            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 0 AS Mes, CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=1 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, ")
            loComandoSeleccionar.AppendLine("           0 As Unidades ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpUnidades ")

            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 1 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=2 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 2 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=3 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 3 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=4 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 4 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=5 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 5 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=6 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 6 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=7 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 7 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=8 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 8 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=9 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 9 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=10 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 10 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=11 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 11 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=12 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 12 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=10 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 13 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=11 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 14 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=12 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 15 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=11 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 16 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=12 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 17 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=11 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 18 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=12 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 19 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=11 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 20 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=12 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 21 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=11 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 22 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=12 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 23 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=11 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 24 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=12 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 25 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=11 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 26 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=12 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 27 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=11 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 28 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=12 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 29 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=12 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 30 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=11 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")
            loComandoSeleccionar.AppendLine(" UNION ALL")											  
            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear, " & lcParametro0Hasta & ") - 31 AS Mes,CASE WHEN DATEPART(DAY," & lcParametro0Hasta & ")>=12 THEN DATEPART(YEAR," & lcParametro0Hasta & ") ELSE DATEPART(YEAR," & lcParametro0Hasta & ")-1 END AS Anio, 0 As Unidades")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(dayofyear,Auditorias.Registro) AS  Mes,")
            loComandoSeleccionar.AppendLine(" 		    DATEPART(YEAR,Auditorias.Registro)  AS  Anio,")
            loComandoSeleccionar.AppendLine(" 		    COUNT(Accion)                       AS  Unidades")
            loComandoSeleccionar.AppendLine(" FROM      Auditorias ")
            loComandoSeleccionar.AppendLine(" WHERE     Auditorias.Registro         Between DATEADD(DAY, -32, " & lcParametro0Hasta & ")")
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Auditorias.Cod_Usu      Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Auditorias.Tabla        Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Auditorias.Opcion       Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Auditorias.Documento    Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Auditorias.Codigo       Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Auditorias.Cod_Emp      Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro6Hasta)


            loComandoSeleccionar.AppendLine(" GROUP BY  DATEPART(YEAR,Auditorias.Registro), DATEPART(dayofyear,Auditorias.Registro) ")
 
            loComandoSeleccionar.AppendLine(" SELECT    #tmpUnidades.Mes                                                        AS Mes,")
            loComandoSeleccionar.AppendLine(" 	        #tmpUnidades.Anio                                                       AS Anio,")
            loComandoSeleccionar.AppendLine(" 	        SUM(#tmpUnidades.Unidades)                                              AS Unidades,")
            loComandoSeleccionar.AppendLine(" 	        ROW_NUMBER() OVER(ORDER BY #tmpUnidades.Anio ASC, #tmpUnidades.Mes ASC) AS Row")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTotalUnidades")
            loComandoSeleccionar.AppendLine(" FROM      #tmpUnidades")
            loComandoSeleccionar.AppendLine(" GROUP BY  Mes, Anio")

            loComandoSeleccionar.AppendLine(" SET       @Sum_Ey             =   ((SELECT SUM(#tmpTotalUnidades.Unidades) FROM #tmpTotalUnidades)/12)")
            loComandoSeleccionar.AppendLine(" SET       @Sum_Ex             =   6.5 ")


            loComandoSeleccionar.AppendLine(" SELECT    ((#tmpTotalUnidades.Row-@sum_Ex)*(#tmpTotalUnidades.Row-@sum_Ey))   AS xEx_yEy,")
            loComandoSeleccionar.AppendLine("           ((#tmpTotalUnidades.Row-@sum_Ex)*(#tmpTotalUnidades.Row-@sum_Ex))   AS xEx_xEx")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTotalUnidades002")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTotalUnidades")

            loComandoSeleccionar.AppendLine(" SET       @Sum_Fac_xEx_yEy    =   (SELECT SUM(#tmpTotalUnidades002.xEx_yEy) FROM #tmpTotalUnidades002)")
            loComandoSeleccionar.AppendLine(" SET       @Sum_Fac_xEx_xEx    =   (SELECT SUM(#tmpTotalUnidades002.xEx_xEx) FROM #tmpTotalUnidades002)")

            loComandoSeleccionar.AppendLine(" SET       @Valor_A            =   (@Sum_Fac_xEx_yEy / @Sum_Fac_xEx_xEx)")
            loComandoSeleccionar.AppendLine(" SET       @Valor_B            =   (@Sum_Ey - (@Valor_A * @Sum_Ex))")

            loComandoSeleccionar.AppendLine(" SELECT    #tmpTotalUnidades.Mes,")
            loComandoSeleccionar.AppendLine(" 	        #tmpTotalUnidades.Anio,")
            loComandoSeleccionar.AppendLine(" 	        #tmpTotalUnidades.Unidades,")
            loComandoSeleccionar.AppendLine(" 	        ((@Valor_A * #tmpTotalUnidades.Row) + @Valor_B)  AS  Tendencia")
            loComandoSeleccionar.AppendLine(" INTO      #tmpTotalUnidades003")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTotalUnidades")

            loComandoSeleccionar.AppendLine(" SELECT    CONVERT(NCHAR(10), DATEADD(DAY, #tmpTotalUnidades003.Mes, (Cast(#tmpTotalUnidades003.Anio As varchar) + '0101')), 103)  AS Dia,")
            loComandoSeleccionar.AppendLine(" 	        #tmpTotalUnidades003.Anio                           AS Anio,")
            loComandoSeleccionar.AppendLine(" 	        #tmpTotalUnidades003.Unidades                       AS Transacciones,")
            loComandoSeleccionar.AppendLine(" 	        #tmpTotalUnidades003.Tendencia                      AS Tendencia")
            loComandoSeleccionar.AppendLine(" FROM      #tmpTotalUnidades003")
            loComandoSeleccionar.AppendLine(" Order By #tmpTotalUnidades003.Mes Asc")
'me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Carga la imagen del logo en cusReportes
            '-------------------------------------------------------------------------------------------------------
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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grTransacciones_32Dias", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrTransacciones_32Dias.ReportSource = loObjetoReporte

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
' CMS: 30/06/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 15/04/11: Ajuste de la vista de Diseño
'-------------------------------------------------------------------------------------------'
