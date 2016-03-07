'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grCobros_MensualesSucursal"
'-------------------------------------------------------------------------------------------'
Partial Class grCobros_MensualesSucursal
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

            If cusAplicacion.goReportes.paParametrosIniciales(0) = 0 Then
                lcParametro0Desde = "'" & Date.Now.Year & "'"
            End If

            Dim loComandoSeleccionar As New StringBuilder()

            '-------------------------------------------------------------------------------------------'
            ' Construccion de la consulta
            '-------------------------------------------------------------------------------------------'

            ' Se crea una tabla temporal con los meses del año
            loComandoSeleccionar.AppendLine(" SELECT 1 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes INTO #tablaFechas")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 2 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 3 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 4 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 5 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 6 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 7 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 8 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 9 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 10 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 11 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT 12 AS Mes," & lcParametro0Desde & " AS Año, 0 As Cobro_Mes")
            loComandoSeleccionar.AppendLine(" ")

            ' Tabla base con todas las sucursales y los meses del año
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine("       Sucursales.Cod_Suc,")
            loComandoSeleccionar.AppendLine(" 		Sucursales.Nom_Suc,")
            loComandoSeleccionar.AppendLine(" 		#tablaFechas.Mes,")
            loComandoSeleccionar.AppendLine(" 		#tablaFechas.Año,")
            loComandoSeleccionar.AppendLine(" 		#tablaFechas.Cobro_Mes")
            loComandoSeleccionar.AppendLine(" INTO	#tablaResVerBase1")
            loComandoSeleccionar.AppendLine(" FROM	Sucursales")
            loComandoSeleccionar.AppendLine(" CROSS JOIN #tablaFechas")
            loComandoSeleccionar.AppendLine(" ")

            ' Tabla resultado con la informacion de los cobros mensuales por sucursal (vertical)
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine(" 		Sucursales.Cod_Suc,")
            loComandoSeleccionar.AppendLine(" 		Sucursales.Nom_Suc,")
            loComandoSeleccionar.AppendLine("   	DATEPART(MONTH, Cobros.fec_ini)AS Mes,")
            loComandoSeleccionar.AppendLine("   	DATEPART(YEAR, Cobros.fec_ini) AS Año,")
            loComandoSeleccionar.AppendLine("   	SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) AS Cobro_Mes")
            loComandoSeleccionar.AppendLine("INTO	#tablaResVerBase2")
            loComandoSeleccionar.AppendLine("FROM	Cobros")
            'loComandoSeleccionar.AppendLine("JOIN	Detalles_Cobros ON Detalles_Cobros.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Cobros ON Renglones_Cobros.Documento = Cobros.Documento")
			loComandoSeleccionar.AppendLine("JOIN	Cuentas_Cobrar AS Cuentas_Cobrar ON Cuentas_Cobrar.Documento = Renglones_Cobros.Doc_ori")
            loComandoSeleccionar.AppendLine(" 		AND	Cuentas_Cobrar.Cod_Tip = Renglones_Cobros.Cod_Tip")
            loComandoSeleccionar.AppendLine("JOIN Sucursales ON  Sucursales.Cod_Suc = Cuentas_Cobrar.Cod_Suc")
            loComandoSeleccionar.AppendLine("WHERE Cobros.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("  		AND DATEPART(YEAR, Cobros.Fec_ini) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("  		AND Cobros.Cod_Cli BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Cobros.Cod_Ven BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            'loComandoSeleccionar.AppendLine("		AND Cobros.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Cobros.Cod_Mon BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("   	AND Cobros.Cod_Rev BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("   	AND Cuentas_Cobrar.Cod_Suc BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY Sucursales.Cod_Suc, Sucursales.Nom_Suc, DATEPART(YEAR, Cobros.fec_ini), DATEPART(MONTH, Cobros.fec_ini)")
            loComandoSeleccionar.AppendLine(" ")

            ' Tabla resultado con la informacion de los cobros mensuales por sucursal (vertical)
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		#tablaResVerBase1.Cod_Suc,")
            loComandoSeleccionar.AppendLine(" 		#tablaResVerBase1.Nom_Suc,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase1.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine(" 		END AS Str_Mes,")
            loComandoSeleccionar.AppendLine("       #tablaResVerBase1.Mes,")
            loComandoSeleccionar.AppendLine("       #tablaResVerBase1.Año,")
            loComandoSeleccionar.AppendLine("       #tablaResVerBase1.Cobro_Mes")
            loComandoSeleccionar.AppendLine(" INTO #tablaResVer")
            loComandoSeleccionar.AppendLine(" FROM #tablaResVerBase1")
            loComandoSeleccionar.AppendLine(" UNION ALL")
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		#tablaResVerBase2.Cod_Suc,")
            loComandoSeleccionar.AppendLine(" 		#tablaResVerBase2.Nom_Suc,")
            loComandoSeleccionar.AppendLine("       CASE")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 1 THEN 'Ene'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 2 THEN 'Feb'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 3 THEN 'Mar'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 4 THEN 'Abr'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 5 THEN 'May'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 6 THEN 'Jun'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 7 THEN 'Jul'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 8 THEN 'Ago'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 9 THEN 'Sep'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 10 THEN 'Oct'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 11 THEN 'Nov'")
            loComandoSeleccionar.AppendLine("           WHEN #tablaResVerBase2.Mes = 12 THEN 'Dic'")
            loComandoSeleccionar.AppendLine(" 		END AS Str_Mes,")
            loComandoSeleccionar.AppendLine("       #tablaResVerBase2.Mes,")
            loComandoSeleccionar.AppendLine("       #tablaResVerBase2.Año,     ")
            loComandoSeleccionar.AppendLine("       #tablaResVerBase2.Cobro_Mes")
            loComandoSeleccionar.AppendLine(" FROM #tablaResVerBase2")
            loComandoSeleccionar.AppendLine(" ")

            'tabla base de resultado con la informacion de los cobros mensuales por sucursal (horizonal)
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		Sucursales.Cod_Suc,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=1) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_Ene,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=2) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_Feb,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=3) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_Mar,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=4) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_Abr,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=5) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_May,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=6) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_Jun,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=7) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_Jul,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=8) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_Ago,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=9) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_Sep,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=10) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_Oct,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=11) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END")
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_Nov,")
            loComandoSeleccionar.AppendLine("       CASE WHEN (DATEPART(MONTH, Cobros.fec_ini)=12) THEN SUM(ISNULL(Renglones_Cobros.Mon_Abo,0)*")
			loComandoSeleccionar.AppendLine("				CASE  Renglones_Cobros.Tip_Doc WHEN 'Credito' THEN -1 WHEN 'Debito' THEN 1 ELSE 0 END") 		            
            loComandoSeleccionar.AppendLine("		) ELSE 0 END AS Monto_Dic")
            loComandoSeleccionar.AppendLine(" INTO #tablaResHorBase1")
            loComandoSeleccionar.AppendLine(" FROM Cobros")
            'loComandoSeleccionar.AppendLine(" JOIN Detalles_Cobros ON Detalles_Cobros.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Cobros ON Renglones_Cobros.Documento = Cobros.Documento")
			loComandoSeleccionar.AppendLine("JOIN	Cuentas_Cobrar AS Cuentas_Cobrar ON Cuentas_Cobrar.Documento = Renglones_Cobros.Doc_ori")
            loComandoSeleccionar.AppendLine(" 		AND	Cuentas_Cobrar.Cod_Tip = Renglones_Cobros.Cod_Tip")
            loComandoSeleccionar.AppendLine(" JOIN Sucursales ON  Sucursales.Cod_Suc = Cuentas_Cobrar.Cod_Suc")
            loComandoSeleccionar.AppendLine(" WHERE Cobros.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("       AND DATEPART(YEAR, Cobros.Fec_ini) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Cli BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Ven BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            'loComandoSeleccionar.AppendLine("       AND Cobros.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Mon BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Rev BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Suc BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY Sucursales.Cod_Suc, Sucursales.Nom_Suc, DATEPART(YEAR, Cobros.fec_ini), DATEPART(MONTH, Cobros.fec_ini)")
            loComandoSeleccionar.AppendLine(" ")

            ' Tabla resultado con la informacion de los cobros mensuales por sucursal (horizontal)
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		#tablaResHorBase1.Cod_Suc,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_Ene) AS Monto_Ene,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_Feb) AS Monto_Feb,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_Mar) AS Monto_Mar,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_Abr) AS Monto_Abr,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_May) AS Monto_May,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_Jun) AS Monto_Jun,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_Jul) AS Monto_Jul,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_Ago) AS Monto_Ago,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_Sep) AS Monto_Sep,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_Oct) AS Monto_Oct,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_Nov) AS Monto_Nov,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResHorBase1.Monto_dic) AS Monto_Dic")
            loComandoSeleccionar.AppendLine(" INTO #tablaResHor")
            loComandoSeleccionar.AppendLine(" FROM #tablaResHorBase1")
            loComandoSeleccionar.AppendLine(" GROUP BY #tablaResHorBase1.Cod_Suc")
            loComandoSeleccionar.AppendLine(" ")

            ' Tabla resultado final con los montos de los cobros mensuales por sucursal
            loComandoSeleccionar.AppendLine(" SELECT")
            loComandoSeleccionar.AppendLine(" 		#tablaResVer.Cod_Suc,")
            loComandoSeleccionar.AppendLine(" 		#tablaResVer.Nom_Suc,")
            loComandoSeleccionar.AppendLine(" 		#tablaResVer.Str_Mes,")
            loComandoSeleccionar.AppendLine("       #tablaResVer.Mes,")
            loComandoSeleccionar.AppendLine("       #tablaResVer.Año,")
            loComandoSeleccionar.AppendLine("       SUM(#tablaResVer.Cobro_Mes) AS Cobro_Mes,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_Ene,0) AS Monto_Ene,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_Feb,0) AS Monto_Feb,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_Mar,0) AS Monto_Mar,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_Abr,0) AS Monto_Abr,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_May,0) AS Monto_May,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_Jun,0) AS Monto_Jun,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_Jul,0) AS Monto_Jul,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_Ago,0) AS Monto_Ago,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_Sep,0) AS Monto_Sep,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_Oct,0) AS Monto_Oct,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_Nov,0) AS Monto_Nov,")
            loComandoSeleccionar.AppendLine(" 		ISNULL(#tablaResHor.Monto_Dic,0) AS Monto_Dic,")
            loComandoSeleccionar.AppendLine(" 		(ISNULL(#tablaResHor.Monto_Ene,0) + ISNULL(#tablaResHor.Monto_Feb,0) + ISNULL(#tablaResHor.Monto_Mar,0) + ")
            loComandoSeleccionar.AppendLine(" 			ISNULL(#tablaResHor.Monto_Abr,0) + ISNULL(#tablaResHor.Monto_May,0) + ISNULL(#tablaResHor.Monto_Jun,0) + ")
            loComandoSeleccionar.AppendLine(" 			ISNULL(#tablaResHor.Monto_Jul,0) + ISNULL(#tablaResHor.Monto_Ago,0) + ISNULL(#tablaResHor.Monto_Sep,0) + ")
            loComandoSeleccionar.AppendLine(" 			ISNULL(#tablaResHor.Monto_Oct,0) + ISNULL(#tablaResHor.Monto_Nov,0) + ISNULL(#tablaResHor.Monto_Dic,0)) AS Tot_Cobro")
            loComandoSeleccionar.AppendLine(" FROM #tablaResVer")
            loComandoSeleccionar.AppendLine(" LEFT OUTER JOIN #tablaResHor ON #tablaResVer.Cod_Suc = #tablaResHor.Cod_Suc")
            loComandoSeleccionar.AppendLine(" GROUP BY #tablaResVer.Cod_Suc, #tablaResVer.Nom_Suc, #tablaResVer.Str_Mes, #tablaResVer.Mes, #tablaResVer.Año,")
            loComandoSeleccionar.AppendLine("       #tablaResHor.Monto_Ene, #tablaResHor.Monto_Feb, #tablaResHor.Monto_Mar, #tablaResHor.Monto_Abr,")
            loComandoSeleccionar.AppendLine("       #tablaResHor.Monto_May, #tablaResHor.Monto_Jun, #tablaResHor.Monto_Jul, #tablaResHor.Monto_Ago,")
            loComandoSeleccionar.AppendLine("       #tablaResHor.Monto_Sep, #tablaResHor.Monto_Oct, #tablaResHor.Monto_Nov, #tablaResHor.Monto_Dic")
            loComandoSeleccionar.AppendLine(" ORDER BY #tablaResVer.Cod_Suc, #tablaResVer.Año, #tablaResVer.Mes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grCobros_MensualesSucursal", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrCobros_MensualesSucursal.ReportSource = loObjetoReporte

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
' DLC: 11/05/2010: Codigo inicial
'-------------------------------------------------------------------------------------------'
' DLC: 03/09/2010: - Si por parametro el año es 0(cero), se toma el año en curso
'                  - Ajuste de la consulta para que tome solo los cobros distintos de Anulado
'-------------------------------------------------------------------------------------------'
' RJG: 09/09/2010: Ajuste al filtro para tomar solo los cobros Confirmados, y				'
'				   agrupar/filtrar por la sucursal de la CxC.								'
'-------------------------------------------------------------------------------------------'
