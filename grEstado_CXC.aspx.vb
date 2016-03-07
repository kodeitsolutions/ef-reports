'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grEstado_CXC"
'-------------------------------------------------------------------------------------------'
Partial Class grEstado_CXC
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosIniciales(6)
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            'Especificos de Análisis del Vencimiento
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosFinales(10)


            Dim Fecha As String

            If lcParametro10Desde = "Vencimiento" Then
                Fecha = "Fec_Fin"
            Else
                Fecha = "Fec_Ini"
            End If

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine(" ---Inicio Cuentas por Cobrar por Vendedor")
            loConsulta.AppendLine(" SELECT ")
            loConsulta.AppendLine(" 			Cuentas_Cobrar.Cod_Ven AS G3_Cod_Ven,")
            loConsulta.AppendLine(" 			Cuentas_Cobrar.Mon_Bru AS G3_Mon_Bru,  ")
            loConsulta.AppendLine(" 			Cuentas_Cobrar.Mon_Imp1 AS G3_Mon_Imp1,  ")
            loConsulta.AppendLine(" 			Cuentas_Cobrar.Mon_Net G3_Mon_Net,  ")
            loConsulta.AppendLine(" 			(CASE WHEN Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal *(-1) ELSE Cuentas_Cobrar.Mon_Sal END) AS G3_Mon_Sal,   ")
            loConsulta.AppendLine(" 			Vendedores.Nom_Ven AS G3_Nom_Ven    ")
            loConsulta.AppendLine(" INTO		#tempCXCVENDEDOR")
            loConsulta.AppendLine(" FROM		Cuentas_Cobrar,  ")
            loConsulta.AppendLine(" 			Vendedores")
            loConsulta.AppendLine(" WHERE		Cuentas_Cobrar.Cod_Ven = Vendedores.Cod_Ven  ")
           loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Tip between " & lcParametro0Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Fec_Ini between " & lcParametro1Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro1Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Cli between " & lcParametro2Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro2Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Ven between " & lcParametro3Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Status IN (" & lcParametro4Desde & ")")

            If lcParametro6Desde = "Igual" Then
                loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev between " & lcParametro5Desde)
            Else
                loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro5Desde)
            End If
            loConsulta.AppendLine(" 				AND " & lcParametro5Hasta)
            loConsulta.AppendLine("               AND Cuentas_Cobrar.Cod_Suc between " & lcParametro7Desde)
            loConsulta.AppendLine("               AND " & lcParametro7Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Mon between " & lcParametro8Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro8Hasta)
            loConsulta.AppendLine("		    AND ((" & lcParametro9Desde & " = 'Si' AND Cuentas_Cobrar.Mon_Sal > 0)")
            loConsulta.AppendLine("			OR (" & lcParametro9Desde & " <> 'Si' AND (Cuentas_Cobrar.Mon_Sal >= 0 or Cuentas_Cobrar.Mon_Sal < 0)))")
            loConsulta.AppendLine(" ORDER BY Cuentas_Cobrar.Mon_Sal DESC ")
            
            loConsulta.AppendLine(" SELECT    TOP 10 ")
            loConsulta.AppendLine("           3 AS Grafico,")
            loConsulta.AppendLine(" 			'FACT'  AS G1_Cod_Tip,")
            loConsulta.AppendLine(" 			''  AS G1_Nom_Tip,")
            loConsulta.AppendLine(" 			0.0 AS G1_Cant_Doc,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Bas1,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Imp1,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Net,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Sal,")
            loConsulta.AppendLine(" 			'30 Días'  AS G2_Dias,")
            loConsulta.AppendLine(" 			0.0 AS G2_Mon_Net,")
            loConsulta.AppendLine(" 			G3_Cod_Ven,")
            loConsulta.AppendLine(" 			G3_Nom_Ven,")
            loConsulta.AppendLine(" 			SUM(G3_Mon_Bru)  AS G3_Mon_Bru,")
            loConsulta.AppendLine(" 			SUM(G3_Mon_Imp1) AS G3_Mon_Imp1,")
            loConsulta.AppendLine(" 			SUM(G3_Mon_Net)  AS G3_Mon_Net,")
           'loComandoSeleccionar.AppendLine(" 			SUM(G3_Mon_Sal)/100  AS G3_Mon_Sal,")
            loConsulta.AppendLine(" 			SUM(G3_Mon_Sal)  AS G3_Mon_Sal,")
            loConsulta.AppendLine(" 			'30 Días'  AS G4_Dias,")
            loConsulta.AppendLine(" 			0.0 AS G4_Mon_Sal")
            loConsulta.AppendLine(" INTO		#tempCXCVENDEDOR2")
            loConsulta.AppendLine(" FROM		#tempCXCVENDEDOR")
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            'loComandoSeleccionar.AppendLine(" WHERE		G3_Mon_Sal > 0")
            loConsulta.AppendLine(" GROUP BY G3_Cod_Ven, G3_Nom_Ven ")
            loConsulta.AppendLine(" ORder BY G3_Mon_Sal DESC, G3_Cod_Ven")
            loConsulta.AppendLine(" ---Fin Cuentas por Cobrar por Vendedor")


            loConsulta.AppendLine(" ---Inicio Cuentas por Cobrar pot Tipo")
            loConsulta.AppendLine(" SELECT	1 AS Grafico, ")
            loConsulta.AppendLine("           Cuentas_Cobrar.cod_tip AS G1_Cod_Tip,  ")
            loConsulta.AppendLine(" 			Tipos_Documentos.nom_tip AS G1_Nom_Tip,  ")
            loConsulta.AppendLine(" 			COUNT (Cuentas_Cobrar.cod_tip) AS G1_Cant_Doc,  ")
            loConsulta.AppendLine(" 			SUM(Cuentas_Cobrar.mon_bas1) AS G1_Mon_Bas1,  ")
            loConsulta.AppendLine(" 			SUM(Cuentas_Cobrar.mon_imp1) AS G1_Mon_Imp1,  ")
            loConsulta.AppendLine(" 			SUM(Cuentas_Cobrar.mon_net) AS G1_Mon_Net,  ")
            'loComandoSeleccionar.AppendLine(" 			SUM(CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' THEN Cuentas_Cobrar.Mon_Sal *(-1) ELSE Cuentas_Cobrar.Mon_Sal END)/100 AS G1_Mon_Sal,")
            loConsulta.AppendLine(" 			SUM(Cuentas_Cobrar.Mon_Sal ) AS G1_Mon_Sal,")
            ' loComandoSeleccionar.AppendLine(" 			SUM(Cuentas_Cobrar.Mon_Sal )/100 AS G1_Mon_Sal,")
            loConsulta.AppendLine(" 			'30 Días'  AS G2_Dias,")
            loConsulta.AppendLine(" 			0.0 AS G2_Mon_Net,")
            loConsulta.AppendLine(" 			(SELECT MAX(G3_Cod_Ven) FROM #tempCXCVENDEDOR2)  AS G3_Cod_Ven,")
            loConsulta.AppendLine(" 			''  AS G3_Nom_ven,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Bru,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Imp1,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Net,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Sal,")
            loConsulta.AppendLine(" 			'30 Días'  AS G4_Dias,")
            loConsulta.AppendLine(" 			0.0 AS G4_Mon_Sal")
            loConsulta.AppendLine(" INTO		#temp_CUENTAS_CXC")
            loConsulta.AppendLine(" FROM		Cuentas_Cobrar,  ")
            loConsulta.AppendLine(" 			Tipos_Documentos,  ")
            loConsulta.AppendLine(" 			Clientes  ")
            loConsulta.AppendLine(" WHERE 		Cuentas_Cobrar.Cod_tip = Tipos_Documentos.Cod_tip ")
            loConsulta.AppendLine(" 			AND Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli ")
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Tip between " & lcParametro0Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Fec_Ini between " & lcParametro1Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro1Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Cli between " & lcParametro2Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro2Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Ven between " & lcParametro3Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Status IN (" & lcParametro4Desde & ")")
            If lcParametro6Desde = "Igual" Then
                loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev between " & lcParametro5Desde)
            Else
                loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro5Desde)
            End If
            loConsulta.AppendLine(" 				AND " & lcParametro5Hasta)
            loConsulta.AppendLine("               AND Cuentas_Cobrar.Cod_Suc between " & lcParametro7Desde)
            loConsulta.AppendLine("               AND " & lcParametro7Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Mon between " & lcParametro8Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro8Hasta)
            loConsulta.AppendLine("             AND       ((" & lcParametro9Desde & " = 'Si' AND Cuentas_Cobrar.Mon_Sal > 0)")
            loConsulta.AppendLine("             OR        (" & lcParametro9Desde & " <> 'Si' AND (Cuentas_Cobrar.Mon_Sal >= 0 or Cuentas_Cobrar.Mon_Sal < 0)))")
            loConsulta.AppendLine(" GROUP BY Cuentas_Cobrar.Cod_tip, Tipos_Documentos.nom_tip")
            loConsulta.AppendLine(" ---Fin Cuentas por Cobrar pot Tipo")

            loConsulta.AppendLine(" ---Inicio Análisis del Vencimiento")
            loConsulta.AppendLine(" SELECT   ")
            loConsulta.AppendLine("  				CASE   ")
            loConsulta.AppendLine("  					WHEN Cuentas_Cobrar." & Fecha & " > " & lcParametro1Hasta & " THEN 'Por Vencer'")
            loConsulta.AppendLine("  					WHEN (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro1Hasta & ") >= 1) AND (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro1Hasta & ") <= 30) THEN '30 Días'")
            loConsulta.AppendLine("  					WHEN (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro1Hasta & ") >= 31) AND (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro1Hasta & ") <= 60) THEN '60 Días'")
            loConsulta.AppendLine("  					WHEN (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro1Hasta & ") >= 61) AND (DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro1Hasta & ") <= 90) THEN '90 Días'")
            loConsulta.AppendLine("  					WHEN DATEDIFF(d, Cuentas_Cobrar." & Fecha & ", " & lcParametro1Hasta & ") >= 91 THEN '90+ Días'  ")
            loConsulta.AppendLine("  				END AS G2_Dias,")
            'loComandoSeleccionar.AppendLine("  				(CASE WHEN Tip_Doc = 'Credito' THEN cuentas_cobrar.Mon_Sal *(-1) ELSE cuentas_cobrar.Mon_Sal END) AS G2_Mon_Net ")
            loConsulta.AppendLine("  				cuentas_cobrar.Mon_Sal  AS G2_Mon_Net ")
            loConsulta.AppendLine("  INTO			#tempVENCIMIENTO 		")
            loConsulta.AppendLine("  FROM			Cuentas_Cobrar   ")
            loConsulta.AppendLine("  WHERE		cuentas_cobrar.Mon_Sal <> 0 ")
            loConsulta.AppendLine(" 				AND cuentas_cobrar.Status <> 'Anulado' ")
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Tip between " & lcParametro0Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Fec_Ini between " & lcParametro1Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro1Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Cli between " & lcParametro2Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro2Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Ven between " & lcParametro3Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Status IN (" & lcParametro4Desde & ")")
            If lcParametro6Desde = "Igual" Then
                loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev between " & lcParametro5Desde)
            Else
                loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro5Desde)
            End If
            loConsulta.AppendLine(" 				AND " & lcParametro5Hasta)
            loConsulta.AppendLine("               AND Cuentas_Cobrar.Cod_Suc between " & lcParametro7Desde)
            loConsulta.AppendLine("               AND " & lcParametro7Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Mon between " & lcParametro8Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro8Hasta)
            loConsulta.AppendLine("               AND             ((" & lcParametro9Desde & " = 'Si' AND (DATEDIFF(d, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") > 1))")
            loConsulta.AppendLine("               OR              (" & lcParametro9Desde & " <> 'Si' AND ((DATEDIFF(d, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") >= 1) or (DATEDIFF(d, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") < 1))))")
            loConsulta.AppendLine(" ORDER BY   Cuentas_Cobrar.Documento, Cuentas_Cobrar.Fec_Ini, Cuentas_Cobrar.Fec_Fin ")
            
            
            loConsulta.AppendLine(" SELECT    2 AS Grafico,")
            loConsulta.AppendLine(" 			'FACT'  AS G1_Cod_Tip,")
            loConsulta.AppendLine(" 			''  AS G1_Nom_Tip,")
            loConsulta.AppendLine(" 			0.0 AS G1_Cant_Doc,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Bas1,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Imp1,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Net,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Sal,")
            loConsulta.AppendLine(" 			G2_Dias,")
            loConsulta.AppendLine(" 			SUM(G2_Mon_Net) AS G2_Mon_Net,")
           ' loComandoSeleccionar.AppendLine(" 			SUM(G2_Mon_Net)/100 AS G2_Mon_Net,")
            loConsulta.AppendLine(" 			(SELECT MAX(G3_Cod_Ven) FROM #tempCXCVENDEDOR2)  AS G3_Cod_Ven,")
            loConsulta.AppendLine(" 			''  AS G3_Nom_ven,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Bru,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Imp1,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Net,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Sal,")
            loConsulta.AppendLine(" 			'30 Días'  AS G4_Dias,")
            loConsulta.AppendLine(" 			0.0 AS G4_Mon_Sal")
            loConsulta.AppendLine(" INTO		#tempVENCIMIENTO2")
            loConsulta.AppendLine(" FROM		#tempVENCIMIENTO")
            loConsulta.AppendLine(" GROUP BY G2_Dias")
            loConsulta.AppendLine(" ORDER BY G2_Dias")
            loConsulta.AppendLine(" ---Fin Análisis del Vencimiento")


            loConsulta.AppendLine(" ---Inicio Estimación de Cobranza")
            loConsulta.AppendLine("  SELECT")
            loConsulta.AppendLine("  CASE   ")
            loConsulta.AppendLine("  		WHEN CAST(Cuentas_Cobrar.Fec_Fin AS DATE) >= CAST(" & lcParametro1Hasta & " AS DATE) THEN 'Por Vencer'")
            loConsulta.AppendLine("  		WHEN (DATEDIFF(DAY, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") >= 1) AND (DATEDIFF(DAY, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") <= 15) THEN '15 Días' ")
            loConsulta.AppendLine("  		WHEN (DATEDIFF(DAY, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") >= 16) AND (DATEDIFF(DAY, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") <= 31) THEN '30 Días'")
            loConsulta.AppendLine("  		WHEN (DATEDIFF(DAY, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") >= 31) AND (DATEDIFF(DAY, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") <= 60) THEN '60 Días'")
            loConsulta.AppendLine("  		WHEN (DATEDIFF(DAY, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") >= 61) AND (DATEDIFF(DAY, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") <= 90) THEN '90 Días'")
            loConsulta.AppendLine("  		WHEN DATEDIFF(DAY, Cuentas_Cobrar.Fec_Fin, " & lcParametro1Hasta & ") >= 91 THEN '90+ Días'")
            loConsulta.AppendLine("  	END AS G4_Dias,")
            loConsulta.AppendLine("   cuentas_cobrar.Mon_Sal AS G4_Mon_Sal")
            loConsulta.AppendLine("  INTO #tempESTIMACIO_COBRANZA")
            loConsulta.AppendLine("  FROM Cuentas_Cobrar")
            loConsulta.AppendLine("  JOIN Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loConsulta.AppendLine("  WHERE     ")
            loConsulta.AppendLine(" 				Cuentas_Cobrar.Cod_Tip between " & lcParametro0Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Fec_Ini between " & lcParametro1Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro1Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Cli between " & lcParametro2Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro2Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Ven between " & lcParametro3Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Status IN (" & lcParametro4Desde & ")")
            If lcParametro6Desde = "Igual" Then
                loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev between " & lcParametro5Desde)
            Else
                loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Rev NOT between " & lcParametro5Desde)
            End If
            loConsulta.AppendLine(" 				AND " & lcParametro5Hasta)
            loConsulta.AppendLine("               AND Cuentas_Cobrar.Cod_Suc between " & lcParametro7Desde)
            loConsulta.AppendLine("               AND " & lcParametro7Hasta)
            loConsulta.AppendLine(" 				AND Cuentas_Cobrar.Cod_Mon between " & lcParametro8Desde)
            loConsulta.AppendLine(" 				AND " & lcParametro8Hasta)
            loConsulta.AppendLine("  ORDER BY Cuentas_Cobrar.Fec_Fin DESC")
            
            
            
            loConsulta.AppendLine("  SELECT   4 AS Grafico,")
            loConsulta.AppendLine(" 			'FACT'  AS G1_Cod_Tip,")
            loConsulta.AppendLine(" 			''  AS G1_Nom_Tip,")
            loConsulta.AppendLine(" 			0.0 AS G1_Cant_Doc,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Bas1,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Imp1,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Net,")
            loConsulta.AppendLine(" 			0.0 AS G1_Mon_Sal,")
            loConsulta.AppendLine(" 			'30 Días'  AS G2_Dias,")
            loConsulta.AppendLine(" 			0.0 AS G2_Mon_Net,")
            loConsulta.AppendLine(" 			(SELECT MAX(G3_Cod_Ven) FROM #tempCXCVENDEDOR2)  AS G3_Cod_Ven,")
            loConsulta.AppendLine(" 			''  AS G3_Nom_ven,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Bru,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Imp1,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Net,")
            loConsulta.AppendLine(" 			0.0 AS G3_Mon_Sal,")
            loConsulta.AppendLine(" 			G4_Dias,")
            loConsulta.AppendLine("  			Sum(G4_Mon_Sal) As G4_Mon_Sal")
            'loComandoSeleccionar.AppendLine("  			Sum(G4_Mon_Sal)/100 As G4_Mon_Sal")
            loConsulta.AppendLine("  INTO	#tempESTIMACIO_COBRANZA2")
            loConsulta.AppendLine("  FROM	#tempESTIMACIO_COBRANZA")
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            'loComandoSeleccionar.AppendLine(" WHERE		G4_Mon_Sal <> 0")
            loConsulta.AppendLine(" GROUP BY G4_Dias")
            loConsulta.AppendLine(" ORDER BY G4_Dias")
            loConsulta.AppendLine(" ---Fin Estimación de Cobranza")


            loConsulta.AppendLine(" SELECT * FROM #temp_CUENTAS_CXC WHERE G1_Mon_Sal <> 0")
            loConsulta.AppendLine(" UNION ALL")
            loConsulta.AppendLine(" SELECT * FROM #tempVENCIMIENTO2 WHERE G2_Mon_Net <> 0")
            loConsulta.AppendLine(" UNION ALL")
            loConsulta.AppendLine(" SELECT * FROM #tempCXCVENDEDOR2 WHERE G3_Mon_Sal <> 0")
            loConsulta.AppendLine(" UNION ALL")
            loConsulta.AppendLine(" SELECT * FROM #tempESTIMACIO_COBRANZA2 WHERE G4_Mon_Sal <> 0")

            
            Dim loServicios As New cusDatos.goDatos
            
           'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grEstado_CXC", laDatosReporte)

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

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrEstado_CXC.ReportSource = loObjetoReporte

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
' CMS: 25/08/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 01/02/11: MAntenimiento del reporte. Ajuste de las gráficas.
'-------------------------------------------------------------------------------------------'
' MAT: 28/04/11: Ajuste del Select
'-------------------------------------------------------------------------------------------'
' RJG: 05/12/14: Ajuste en formato de los gráficos. Ajuste en el SELECT (caso de facturas   '
'                "Por vencer".                                                              '
'-------------------------------------------------------------------------------------------'
