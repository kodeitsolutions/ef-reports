﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rDetalle_Condominio003_CCMV_2"
'-------------------------------------------------------------------------------------------'
Partial Class rDetalle_Condominio003_CCMV_2
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

            Dim lnPorcentaje1 As Decimal = CDec(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lnPorcentaje2 As Decimal = CDec(cusAplicacion.goReportes.paParametrosIniciales(5))

            If (lnPorcentaje1 <= 0) Then lnPorcentaje1 = 8
            If (lnPorcentaje2 <= 0) Then lnPorcentaje2 = 7

            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(lnPorcentaje1)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(lnPorcentaje2)

            'Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosIniciales(6).Trim()
            'Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosIniciales(7).Trim()

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Renglones_oPagos .*, ")
            loComandoSeleccionar.AppendLine("		Ordenes_Pagos.Fec_Ini As Fecha, ")
            loComandoSeleccionar.AppendLine("		Conceptos.Nom_Con, Conceptos.Grupo ")
            loComandoSeleccionar.AppendLine("INTO	#tempOrdenes")
            loComandoSeleccionar.AppendLine("FROM	Ordenes_Pagos ")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_oPagos ON Renglones_oPagos.Documento = Ordenes_pagos.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Conceptos on Conceptos.Cod_Con = Renglones_oPagos.Cod_Con")
            loComandoSeleccionar.AppendLine("WHERE	Ordenes_Pagos.Cod_Rev = '02'")
            loComandoSeleccionar.AppendLine("    AND Ordenes_Pagos.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("    AND Conceptos.Clase <> 'NO'")
            loComandoSeleccionar.AppendLine("    AND conceptos.grupo IN ('recurrente', 'NoRecurrente')")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  Articulos.Cod_Cla, ")
            loComandoSeleccionar.AppendLine(" 		Precios_Clientes.Cod_Art AS Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 		Precios_Clientes.Cod_Reg AS Cod_Ori, ")
            loComandoSeleccionar.AppendLine(" 		Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Por_Ali, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Por_Ali2, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Por_Ali3, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Precio1 AS Canon, ")
            loComandoSeleccionar.AppendLine(" 		Articulos.Modelo AS Condominio_Adicional, ")
            loComandoSeleccionar.AppendLine(" 		COALESCE(Campos_Propiedades.Val_Num, 0) AS Met_Cua ")
            loComandoSeleccionar.AppendLine("INTO	#tempClientesArticulos001 ")
            loComandoSeleccionar.AppendLine("FROM	Articulos ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Campos_Propiedades ")
            loComandoSeleccionar.AppendLine("		ON Articulos.Cod_Art = Campos_Propiedades.Cod_Reg ")
            loComandoSeleccionar.AppendLine("		AND Campos_Propiedades.Cod_Pro = 'ART_001' ")
            loComandoSeleccionar.AppendLine("		AND Campos_Propiedades.Origen = 'Articulos' ")
            loComandoSeleccionar.AppendLine("	JOIN Precios_Clientes			")
            loComandoSeleccionar.AppendLine("		ON Articulos.Cod_Art = Precios_Clientes.Cod_Art ")
            loComandoSeleccionar.AppendLine("	JOIN Clientes				 ")
            loComandoSeleccionar.AppendLine("		ON Precios_Clientes.Cod_Reg = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("WHERE   Clientes.Status = 'A' ")
            loComandoSeleccionar.AppendLine("	AND Articulos.Status = 'A' ")
            loComandoSeleccionar.AppendLine("   AND Clientes.Cod_Cli BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("   AND " & lcParametro0Hasta)
            'loComandoSeleccionar.AppendLine("   AND Articulos.Precio1>0")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  #tempClientesArticulos001.Cod_Cla, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos001.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos001.Cod_Ori, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos001.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos001.Por_Ali, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos001.Por_Ali2, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos001.Por_Ali3, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos001.Cod_Art, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos001.Met_Cua, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos001.Condominio_Adicional, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos001.Canon, ")
            loComandoSeleccionar.AppendLine(" 		COALESCE(Campos_Propiedades.Val_Num, 0) AS Pre_Met ")
            loComandoSeleccionar.AppendLine("INTO	#tempClientesArticulos ")
            loComandoSeleccionar.AppendLine("FROM    #tempClientesArticulos001 ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Campos_Propiedades ")
            loComandoSeleccionar.AppendLine("		ON	#tempClientesArticulos001.Cod_Art = Campos_Propiedades.Cod_Reg ")
            loComandoSeleccionar.AppendLine("		AND Campos_Propiedades.Cod_Pro = 'ART_002' ")
            loComandoSeleccionar.AppendLine("		AND Campos_Propiedades.Origen = 'Articulos' ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	#tempClientesArticulos.Cod_Ori, ")
            loComandoSeleccionar.AppendLine("        #tempClientesArticulos.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Por_Ali,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Por_Ali2,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Por_Ali3,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Cod_Art,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Met_Cua,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Pre_Met,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Condominio_Adicional,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Canon As Canon,")
            loComandoSeleccionar.AppendLine(" 		SUM(#tempClientesArticulos.Met_Cua * #tempClientesArticulos.Pre_Met) AS Pre_Loc,")
            loComandoSeleccionar.AppendLine(" 		#tempOrdenes.Grupo          As  Grupo,")
            loComandoSeleccionar.AppendLine(" 		SUM(#tempOrdenes.Mon_Net) As Mon_Net,")
            loComandoSeleccionar.AppendLine(" 		CAST(#tempOrdenes.Nom_Con As Varchar(5000)) As Motivo,")
            loComandoSeleccionar.AppendLine("        DATEPART(YEAR, #tempOrdenes.Fecha) As Anio, ")
            loComandoSeleccionar.AppendLine(" 		DATEPART(MONTH, #tempOrdenes.Fecha) As Mes,")
            loComandoSeleccionar.AppendLine(" 		#tempOrdenes.Cod_Con,")
            loComandoSeleccionar.AppendLine("        SUM(CASE ")
            loComandoSeleccionar.AppendLine("     	    WHEN #tempOrdenes.Cod_Con IN ('E024','E025') THEN ")
            loComandoSeleccionar.AppendLine("     	        #tempOrdenes.Mon_Net  * (#tempClientesArticulos.Por_Ali2/100)")
            loComandoSeleccionar.AppendLine("     	    ELSE")
            loComandoSeleccionar.AppendLine("     	        #tempOrdenes.Mon_Net  * (#tempClientesArticulos.Por_Ali/100)")
            loComandoSeleccionar.AppendLine("			END) AS NetxAli,")
            loComandoSeleccionar.AppendLine("		SUM(CASE ")
            loComandoSeleccionar.AppendLine(" 			WHEN #tempOrdenes.Cod_Con IN ('E024','E025') THEN ")
            loComandoSeleccionar.AppendLine(" 		        0.0")
            loComandoSeleccionar.AppendLine(" 	        ELSE")
            loComandoSeleccionar.AppendLine(" 		        #tempOrdenes.Mon_Net ")
            loComandoSeleccionar.AppendLine("           END) AS NetxAliPlus,")
            loComandoSeleccionar.AppendLine("        SUM(CASE   ")
            loComandoSeleccionar.AppendLine("       	    WHEN #tempOrdenes.Cod_Con IN ('E024','E025') THEN   ")
            loComandoSeleccionar.AppendLine("       		    0  ")
            loComandoSeleccionar.AppendLine("       	    ELSE  ")
            loComandoSeleccionar.AppendLine("       		    #tempOrdenes.Mon_Net  * " & lcParametro4Desde & "  ")
            loComandoSeleccionar.AppendLine("            END) AS NetxAliRes,  ")
            loComandoSeleccionar.AppendLine("        CASE ")
            loComandoSeleccionar.AppendLine("       		WHEN #tempOrdenes.Cod_Con IN ('E024','E025') THEN ")
            loComandoSeleccionar.AppendLine("       			2")
            loComandoSeleccionar.AppendLine("       		ELSE")
            loComandoSeleccionar.AppendLine("       			1")
            loComandoSeleccionar.AppendLine("			END AS Alicuota")
            loComandoSeleccionar.AppendLine("INTO      #temFinal")
            loComandoSeleccionar.AppendLine("FROM      #tempClientesArticulos, #tempOrdenes")
            loComandoSeleccionar.AppendLine("WHERE	 ")
            'loComandoSeleccionar.AppendLine("           Cod_Cli                         BETWEEN " & lcParametro0Desde)
            'loComandoSeleccionar.AppendLine("       AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       #tempOrdenes.Fecha              BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND #tempClientesArticulos.Cod_Cla  BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	DATEPART(YEAR, #tempOrdenes.Fecha), ")
            loComandoSeleccionar.AppendLine(" 			DATEPART(MONTH, #tempOrdenes.Fecha),")
            loComandoSeleccionar.AppendLine(" 		#tempOrdenes.Grupo, ")
            loComandoSeleccionar.AppendLine("		#tempClientesArticulos.Cod_Ori, ")
            loComandoSeleccionar.AppendLine("		#tempClientesArticulos.	Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.	Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.	Por_Ali,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.	Por_Ali2,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.	Por_Ali3,")
            loComandoSeleccionar.AppendLine(" 		CAST(#tempOrdenes.Nom_Con AS VARCHAR(5000)),")
            loComandoSeleccionar.AppendLine(" 		#tempOrdenes.Cod_Con,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Cod_Art,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Met_Cua,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Canon,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Condominio_Adicional,")
            loComandoSeleccionar.AppendLine(" 		#tempClientesArticulos.Pre_Met")
            loComandoSeleccionar.AppendLine("ORDER BY DATEPART(YEAR, #tempOrdenes.Fecha), ")
            loComandoSeleccionar.AppendLine("	DATEPART(MONTH, #tempOrdenes.Fecha), ")
            loComandoSeleccionar.AppendLine("	Cod_Ori, Cod_Cli, Nom_Cli, CAST(#tempOrdenes.Nom_Con As Varchar(5000))")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT  SUBSTRING(Cod_Cli, 1, 15) + ")
            loComandoSeleccionar.AppendLine("       	CASE WHEN dbo.Concatena(Cod_Ori) = '' THEN Cod_Ori")
            loComandoSeleccionar.AppendLine("       		ELSE dbo.Concatena(Cod_Ori)")
            loComandoSeleccionar.AppendLine("       		END AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("       	Nom_Cli,")
            loComandoSeleccionar.AppendLine("       	Por_Ali,")
            loComandoSeleccionar.AppendLine("       	Por_Ali2,")
            loComandoSeleccionar.AppendLine("       	Por_Ali3,")
            loComandoSeleccionar.AppendLine("       	Cod_Art,")
            loComandoSeleccionar.AppendLine("       	Met_Cua,")
            loComandoSeleccionar.AppendLine("       	Pre_Met,")
            loComandoSeleccionar.AppendLine("       	Condominio_Adicional,")
            loComandoSeleccionar.AppendLine("       	Canon,")
            loComandoSeleccionar.AppendLine("       	Pre_Loc,")
            loComandoSeleccionar.AppendLine("       	Mon_Net,")
            loComandoSeleccionar.AppendLine("       	Grupo,")
            loComandoSeleccionar.AppendLine("       	UPPER(Motivo) As Motivo,")
            loComandoSeleccionar.AppendLine("       	Anio,")
            loComandoSeleccionar.AppendLine("       	Mes,")
            loComandoSeleccionar.AppendLine("       	Cod_Con,")
            loComandoSeleccionar.AppendLine("       	NetxAli,")
            loComandoSeleccionar.AppendLine("       	NetxAliPlus, NetxAliRes,")
            loComandoSeleccionar.AppendLine("")
            If lcParametro3Desde = "''" Then
                loComandoSeleccionar.AppendLine("       (SELECT Por_Imp1 FROM Impuestos WHERE Cod_Imp = 'EXE') As Por_Imp1,")
            Else
                loComandoSeleccionar.AppendLine("       (SELECT Por_Imp1 FROM Impuestos WHERE Cod_Imp = " & lcParametro3Desde & ") As Por_Imp1,")
            End If
            loComandoSeleccionar.AppendLine("       	Alicuota,")
            loComandoSeleccionar.AppendLine("       	CAST(" & lcParametro4Desde & " AS DECIMAL(28,10)) AS Porcentaje1,")
            loComandoSeleccionar.AppendLine("       	CAST(" & lcParametro5Desde & " AS DECIMAL(28,10))  AS Porcentaje2, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS INT)               AS L1_Linea, ")
            loComandoSeleccionar.AppendLine("       	SPACE(30)                       AS L1_Cod_Cue, ")
            loComandoSeleccionar.AppendLine("       	SPACE(100)                      AS L1_Nom_Cue, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L1_Mon_Ini, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L1_Mon_Mov, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L1_Mon_Fin, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS INT)               AS L2_Linea, ")
            loComandoSeleccionar.AppendLine("       	SPACE(30)                       AS L2_Cod_Cue, ")
            loComandoSeleccionar.AppendLine("       	SPACE(100)                      AS L2_Nom_Cue, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L2_Mon_Ini, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L2_Mon_Mov, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L2_Mon_Fin, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS INT)               AS L3_Linea, ")
            loComandoSeleccionar.AppendLine("       	SPACE(30)                       AS L3_Cod_Cue, ")
            loComandoSeleccionar.AppendLine("       	SPACE(100)                      AS L3_Nom_Cue, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L3_Mon_Ini, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L3_Mon_Mov, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L3_Mon_Fin, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS INT)               AS L4_Linea, ")
            loComandoSeleccionar.AppendLine("       	SPACE(30)                       AS L4_Cod_Cue, ")
            loComandoSeleccionar.AppendLine("       	SPACE(100)                      AS L4_Nom_Cue, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L4_Mon_Ini, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L4_Mon_Mov, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L4_Mon_Fin, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS INT)               AS L5_Linea, ")
            loComandoSeleccionar.AppendLine("       	SPACE(30)                       AS L5_Cod_Cue, ")
            loComandoSeleccionar.AppendLine("       	SPACE(100)                      AS L5_Nom_Cue, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L5_Mon_Ini, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L5_Mon_Mov, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L5_Mon_Fin, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS INT)               AS L6_Linea, ")
            loComandoSeleccionar.AppendLine("       	SPACE(30)                       AS L6_Cod_Cue, ")
            loComandoSeleccionar.AppendLine("       	SPACE(100)                      AS L6_Nom_Cue, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L6_Mon_Ini, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L6_Mon_Mov, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L6_Mon_Fin, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS INT)               AS L7_Linea, ")
            loComandoSeleccionar.AppendLine("       	SPACE(30)                       AS L7_Cod_Cue, ")
            loComandoSeleccionar.AppendLine("       	SPACE(100)                      AS L7_Nom_Cue, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L7_Mon_Ini, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L7_Mon_Mov, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L7_Mon_Fin, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS INT)               AS L8_Linea, ")
            loComandoSeleccionar.AppendLine("       	SPACE(30)                       AS L8_Cod_Cue, ")
            loComandoSeleccionar.AppendLine("       	SPACE(100)                      AS L8_Nom_Cue, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L8_Mon_Ini, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L8_Mon_Mov, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L8_Mon_Fin, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS INT)               AS L9_Linea, ")
            loComandoSeleccionar.AppendLine("       	SPACE(30)                       AS L9_Cod_Cue, ")
            loComandoSeleccionar.AppendLine("       	SPACE(100)                      AS L9_Nom_Cue, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L9_Mon_Ini, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L9_Mon_Mov, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L9_Mon_Fin, ")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS INT)               AS L10_Linea, ")
            loComandoSeleccionar.AppendLine("       	SPACE(30)                       AS L10_Cod_Cue,")
            loComandoSeleccionar.AppendLine("       	SPACE(100)                      AS L10_Nom_Cue,")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L10_Mon_Ini,")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L10_Mon_Mov,")
            loComandoSeleccionar.AppendLine("       	CAST(0.00 AS DECIMAL(28,10))    AS L10_Mon_Fin")
            loComandoSeleccionar.AppendLine("INTO      #tmpDatosAdministrativos ")
            loComandoSeleccionar.AppendLine("FROM      #temFinal  ")
            loComandoSeleccionar.AppendLine("WHERE     dbo.Concatena(Cod_Ori) <> '' ")
            loComandoSeleccionar.AppendLine("ORDER BY  Anio, Mes, Cod_Ori, ")
            loComandoSeleccionar.AppendLine("		  Cod_Cli, Cod_Con, Alicuota, Grupo DESC")
            loComandoSeleccionar.AppendLine("")

            '-------------------------------------------------------------------------------------------'
            ' Busqueda de los montos iniciales de las cuentas de reservas en la Contabilidad
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Cuentas_Contables.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("		Cuentas_Contables.Nom_Cue, ")
            loComandoSeleccionar.AppendLine("		Iniciales.mon_ini, ")
            loComandoSeleccionar.AppendLine("		COALESCE(SUM(MO.mon_deb) - SUM(MO.mon_hab), 0) AS Movimiento,")
            loComandoSeleccionar.AppendLine("		YEAR(" & lcParametro1Desde & ") AS Anio,")
            loComandoSeleccionar.AppendLine("		MONTH(" & lcParametro1Desde & ") AS Mes ")
            loComandoSeleccionar.AppendLine("INTO	#tmpReservas")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT  RC.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("					SUM(CASE WHEN RC.Fec_Ini < " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("						THEN RC.mon_deb - RC.mon_hab ELSE 0 ")
            loComandoSeleccionar.AppendLine("					END) AS mon_ini")
            loComandoSeleccionar.AppendLine("			FROM	[Factory_Contabilidad_ResCon].[dbo].[Renglones_Comprobantes] RC")
            loComandoSeleccionar.AppendLine("			WHERE	RC.Fec_Ini <= " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("				AND SUBSTRING(RC.Cod_Cue,1,2) = '2.'")
            loComandoSeleccionar.AppendLine("			GROUP BY RC.Cod_Cue")
            loComandoSeleccionar.AppendLine("		) AS Iniciales")
            loComandoSeleccionar.AppendLine("	JOIN [Factory_Contabilidad_ResCon].[dbo].[Cuentas_Contables] ")
            loComandoSeleccionar.AppendLine("		ON Cuentas_Contables.cod_cue = Iniciales.cod_cue")
            loComandoSeleccionar.AppendLine("	LEFT JOIN [Factory_Contabilidad_ResCon].[dbo].[Renglones_Comprobantes] MO")
            loComandoSeleccionar.AppendLine("		ON MO.cod_cue = Iniciales.cod_cue")
            loComandoSeleccionar.AppendLine("		AND MO.fec_ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("                   AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY Cuentas_Contables.Cod_Cue, Cuentas_Contables.Nom_Cue, Iniciales.mon_ini;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	#tmpReservas.Anio,")
            loComandoSeleccionar.AppendLine("		#tmpReservas.Mes,")
            loComandoSeleccionar.AppendLine("		#tmpReservas.Cod_Cue,")
            loComandoSeleccionar.AppendLine("		#tmpReservas.Nom_Cue,")
            loComandoSeleccionar.AppendLine("		-1*#tmpReservas.Mon_Ini AS Mon_Ini,")
            loComandoSeleccionar.AppendLine("		-1*#tmpReservas.Movimiento As Mon_Mov,")
            loComandoSeleccionar.AppendLine("		-1*(#tmpReservas.Mon_Ini + #tmpReservas.Movimiento) As Mon_Fin ")
            loComandoSeleccionar.AppendLine("INTO	#tmpRenglonesComprobantes")
            loComandoSeleccionar.AppendLine("FROM	#tmpReservas")
            loComandoSeleccionar.AppendLine("")


            '-------------------------------------------------------------------------------------------'
            ' Actualizacion de los datos del Select de Trabajo del Administrativo
            ' Con los datos que me traigo de los saldos de las Reservas desde el Sistema Contable
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" DECLARE curTrabajo CURSOR SCROLL KEYSET FOR ")
            loComandoSeleccionar.AppendLine("       SELECT * FROM #tmpRenglonesComprobantes ")
            loComandoSeleccionar.AppendLine("       ")

            ' Se abre el cursor y declaran variables locales
            loComandoSeleccionar.AppendLine(" OPEN curTrabajo ")
            loComandoSeleccionar.AppendLine(" DECLARE @lnAnio       INT ")
            loComandoSeleccionar.AppendLine(" DECLARE @lnMes        INT ")
            loComandoSeleccionar.AppendLine(" DECLARE @lnPosicion   INT ")
            loComandoSeleccionar.AppendLine(" DECLARE @lnTotal      INT ")
            loComandoSeleccionar.AppendLine(" DECLARE @lcCodCue     CHAR(30) ")
            loComandoSeleccionar.AppendLine(" DECLARE @lcNomCue     CHAR(100) ")
            loComandoSeleccionar.AppendLine(" DECLARE @lnMonIni     DECIMAL(28,10) ")
            loComandoSeleccionar.AppendLine(" DECLARE @lnMonMov     DECIMAL(28,10) ")
            loComandoSeleccionar.AppendLine(" DECLARE @lnMonFin     DECIMAL(28,10) ")
            loComandoSeleccionar.AppendLine(" SET @lnPosicion = 1 ")
            loComandoSeleccionar.AppendLine(" SET @lnTotal = (SELECT COUNT(*) FROM #tmpRenglonesComprobantes) ")

            '-------------------------------------------------------------------------------------------'
            ' Inicio del conteo de los renglones
            '-------------------------------------------------------------------------------------------'
            'FETCH para posicionar en primer registro del cursor 
            loComandoSeleccionar.AppendLine(" FETCH NEXT FROM curTrabajo INTO @lnAnio, @lnMes, @lcCodCue, @lcNomCue, @lnMonIni, @lnMonMov, @lnMonFin ")
            loComandoSeleccionar.AppendLine(" WHILE @@FETCH_STATUS = 0 ")
            loComandoSeleccionar.AppendLine(" BEGIN ")
            loComandoSeleccionar.AppendLine("   IF (@lnPosicion = 1) ")
            loComandoSeleccionar.AppendLine("   BEGIN ")
            loComandoSeleccionar.AppendLine("       UPDATE  #tmpDatosAdministrativos ")
            loComandoSeleccionar.AppendLine("       SET     L1_Linea    =    @lnPosicion, ")
            loComandoSeleccionar.AppendLine("               L1_Cod_Cue  =    @lcCodCue, ")
            loComandoSeleccionar.AppendLine("               L1_Nom_Cue  =    @lcNomCue, ")
            loComandoSeleccionar.AppendLine("               L1_Mon_Ini  =    @lnMonIni, ")
            loComandoSeleccionar.AppendLine("               L1_Mon_Mov  =    @lnMonMov, ")
            loComandoSeleccionar.AppendLine("               L1_Mon_Fin  =    @lnMonFin ")
            loComandoSeleccionar.AppendLine("   END ")
            loComandoSeleccionar.AppendLine("   IF (@lnPosicion = 2) ")
            loComandoSeleccionar.AppendLine("   BEGIN ")
            loComandoSeleccionar.AppendLine("       UPDATE  #tmpDatosAdministrativos ")
            loComandoSeleccionar.AppendLine("       SET     L2_Linea    =    @lnPosicion, ")
            loComandoSeleccionar.AppendLine("               L2_Cod_Cue  =    @lcCodCue, ")
            loComandoSeleccionar.AppendLine("               L2_Nom_Cue  =    @lcNomCue, ")
            loComandoSeleccionar.AppendLine("               L2_Mon_Ini  =    @lnMonIni, ")
            loComandoSeleccionar.AppendLine("               L2_Mon_Mov  =    @lnMonMov, ")
            loComandoSeleccionar.AppendLine("               L2_Mon_Fin  =    @lnMonFin ")
            loComandoSeleccionar.AppendLine("   END ")
            loComandoSeleccionar.AppendLine("   IF (@lnPosicion = 3) ")
            loComandoSeleccionar.AppendLine("   BEGIN ")
            loComandoSeleccionar.AppendLine("       UPDATE  #tmpDatosAdministrativos ")
            loComandoSeleccionar.AppendLine("       SET     L3_Linea    =    @lnPosicion, ")
            loComandoSeleccionar.AppendLine("               L3_Cod_Cue  =    @lcCodCue, ")
            loComandoSeleccionar.AppendLine("               L3_Nom_Cue  =    @lcNomCue, ")
            loComandoSeleccionar.AppendLine("               L3_Mon_Ini  =    @lnMonIni, ")
            loComandoSeleccionar.AppendLine("               L3_Mon_Mov  =    @lnMonMov, ")
            loComandoSeleccionar.AppendLine("               L3_Mon_Fin  =    @lnMonFin ")
            loComandoSeleccionar.AppendLine("   END ")
            loComandoSeleccionar.AppendLine("   IF (@lnPosicion = 4) ")
            loComandoSeleccionar.AppendLine("   BEGIN ")
            loComandoSeleccionar.AppendLine("       UPDATE  #tmpDatosAdministrativos ")
            loComandoSeleccionar.AppendLine("       SET     L4_Linea    =    @lnPosicion, ")
            loComandoSeleccionar.AppendLine("               L4_Cod_Cue  =    @lcCodCue, ")
            loComandoSeleccionar.AppendLine("               L4_Nom_Cue  =    @lcNomCue, ")
            loComandoSeleccionar.AppendLine("               L4_Mon_Ini  =    @lnMonIni, ")
            loComandoSeleccionar.AppendLine("               L4_Mon_Mov  =    @lnMonMov, ")
            loComandoSeleccionar.AppendLine("               L4_Mon_Fin  =    @lnMonFin ")
            loComandoSeleccionar.AppendLine("   END ")
            loComandoSeleccionar.AppendLine("   IF (@lnPosicion = 5) ")
            loComandoSeleccionar.AppendLine("   BEGIN ")
            loComandoSeleccionar.AppendLine("       UPDATE  #tmpDatosAdministrativos ")
            loComandoSeleccionar.AppendLine("       SET     L5_Linea    =    @lnPosicion, ")
            loComandoSeleccionar.AppendLine("               L5_Cod_Cue  =    @lcCodCue, ")
            loComandoSeleccionar.AppendLine("               L5_Nom_Cue  =    @lcNomCue, ")
            loComandoSeleccionar.AppendLine("               L5_Mon_Ini  =    @lnMonIni, ")
            loComandoSeleccionar.AppendLine("               L5_Mon_Mov  =    @lnMonMov, ")
            loComandoSeleccionar.AppendLine("               L5_Mon_Fin  =    @lnMonFin ")
            loComandoSeleccionar.AppendLine("   END ")
            loComandoSeleccionar.AppendLine("   IF (@lnPosicion = 6) ")
            loComandoSeleccionar.AppendLine("   BEGIN ")
            loComandoSeleccionar.AppendLine("       UPDATE  #tmpDatosAdministrativos ")
            loComandoSeleccionar.AppendLine("       SET     L6_Linea    =    @lnPosicion, ")
            loComandoSeleccionar.AppendLine("               L6_Cod_Cue  =    @lcCodCue, ")
            loComandoSeleccionar.AppendLine("               L6_Nom_Cue  =    @lcNomCue, ")
            loComandoSeleccionar.AppendLine("               L6_Mon_Ini  =    @lnMonIni, ")
            loComandoSeleccionar.AppendLine("               L6_Mon_Mov  =    @lnMonMov, ")
            loComandoSeleccionar.AppendLine("               L6_Mon_Fin  =    @lnMonFin ")
            loComandoSeleccionar.AppendLine("   END ")
            loComandoSeleccionar.AppendLine("   IF (@lnPosicion = 7) ")
            loComandoSeleccionar.AppendLine("   BEGIN ")
            loComandoSeleccionar.AppendLine("       UPDATE  #tmpDatosAdministrativos ")
            loComandoSeleccionar.AppendLine("       SET     L7_Linea    =    @lnPosicion, ")
            loComandoSeleccionar.AppendLine("               L7_Cod_Cue  =    @lcCodCue, ")
            loComandoSeleccionar.AppendLine("               L7_Nom_Cue  =    @lcNomCue, ")
            loComandoSeleccionar.AppendLine("               L7_Mon_Ini  =    @lnMonIni, ")
            loComandoSeleccionar.AppendLine("               L7_Mon_Mov  =    @lnMonMov, ")
            loComandoSeleccionar.AppendLine("               L7_Mon_Fin  =    @lnMonFin ")
            loComandoSeleccionar.AppendLine("   END ")
            loComandoSeleccionar.AppendLine("   IF (@lnPosicion = 8) ")
            loComandoSeleccionar.AppendLine("   BEGIN ")
            loComandoSeleccionar.AppendLine("       UPDATE  #tmpDatosAdministrativos ")
            loComandoSeleccionar.AppendLine("       SET     L8_Linea    =    @lnPosicion, ")
            loComandoSeleccionar.AppendLine("               L8_Cod_Cue  =    @lcCodCue, ")
            loComandoSeleccionar.AppendLine("               L8_Nom_Cue  =    @lcNomCue, ")
            loComandoSeleccionar.AppendLine("               L8_Mon_Ini  =    @lnMonIni, ")
            loComandoSeleccionar.AppendLine("               L8_Mon_Mov  =    @lnMonMov, ")
            loComandoSeleccionar.AppendLine("               L8_Mon_Fin  =    @lnMonFin ")
            loComandoSeleccionar.AppendLine("   END ")
            loComandoSeleccionar.AppendLine("   IF (@lnPosicion = 9) ")
            loComandoSeleccionar.AppendLine("   BEGIN ")
            loComandoSeleccionar.AppendLine("       UPDATE  #tmpDatosAdministrativos ")
            loComandoSeleccionar.AppendLine("       SET     L9_Linea    =    @lnPosicion, ")
            loComandoSeleccionar.AppendLine("               L9_Cod_Cue  =    @lcCodCue, ")
            loComandoSeleccionar.AppendLine("               L9_Nom_Cue  =    @lcNomCue, ")
            loComandoSeleccionar.AppendLine("               L9_Mon_Ini  =    @lnMonIni, ")
            loComandoSeleccionar.AppendLine("               L9_Mon_Mov  =    @lnMonMov, ")
            loComandoSeleccionar.AppendLine("               L9_Mon_Fin  =    @lnMonFin ")
            loComandoSeleccionar.AppendLine("   END ")
            loComandoSeleccionar.AppendLine("   IF (@lnPosicion = 10) ")
            loComandoSeleccionar.AppendLine("   BEGIN ")
            loComandoSeleccionar.AppendLine("       UPDATE  #tmpDatosAdministrativos ")
            loComandoSeleccionar.AppendLine("       SET     L10_Linea    =    @lnPosicion, ")
            loComandoSeleccionar.AppendLine("               L10_Cod_Cue  =    @lcCodCue, ")
            loComandoSeleccionar.AppendLine("               L10_Nom_Cue  =    @lcNomCue, ")
            loComandoSeleccionar.AppendLine("               L10_Mon_Ini  =    @lnMonIni, ")
            loComandoSeleccionar.AppendLine("               L10_Mon_Mov  =    @lnMonMov, ")
            loComandoSeleccionar.AppendLine("               L10_Mon_Fin  =    @lnMonFin ")
            loComandoSeleccionar.AppendLine("   END ")
            loComandoSeleccionar.AppendLine("   SET @lnPosicion = @lnPosicion + 1 ")
            loComandoSeleccionar.AppendLine("   FETCH NEXT FROM curTrabajo INTO @lnAnio, @lnMes, @lcCodCue, @lcNomCue, @lnMonIni, @lnMonMov, @lnMonFin ")
            loComandoSeleccionar.AppendLine(" END ")
            loComandoSeleccionar.AppendLine("")

            'Se cierra el cursor
            loComandoSeleccionar.AppendLine(" CLOSE curTrabajo ")
            loComandoSeleccionar.AppendLine(" DEALLOCATE curTrabajo ")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT *, ")
            loComandoSeleccionar.AppendLine("       REPLACE(CAST(CAST(por_ali  AS DECIMAL(28,14)) AS VARCHAR(20)), '.', ',') AS Alicuota_1_Completa,")
            loComandoSeleccionar.AppendLine("       REPLACE(CAST(CAST(por_ali2 AS DECIMAL(28,14)) AS VARCHAR(20)), '.', ',') AS Alicuota_2_Completa")
            loComandoSeleccionar.AppendLine("FROM #tmpDatosAdministrativos")
            'If (lcParametro6Desde.ToUpper() = "SI") Then
            '    loComandoSeleccionar.AppendLine("WHERE Canon > 0")
            'ElseIf (lcParametro7Desde.ToUpper() = "SI") Then
            '    loComandoSeleccionar.AppendLine("WHERE Canon <= 0")
            'End If
            loComandoSeleccionar.AppendLine("ORDER BY  Anio, Mes, ")
            loComandoSeleccionar.AppendLine("		  Cod_Cli, Alicuota, Grupo DESC, Cod_Con  ")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rDetalle_Condominio003_CCMV_2", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrDetalle_Condominio003_CCMV_2.ReportSource = loObjetoReporte

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
' CMS: 02/07/10: Codigo inicial
'-------------------------------------------------------------------------------------------'
' JJD: 20/07/10: Ajustes para generar un recibo por cada cliente asociado
'-------------------------------------------------------------------------------------------'
' JFP: 04/11/11: Filtrado de los conceptos con campo clase <> 'NO'
'-------------------------------------------------------------------------------------------'
' JJD: 05/02/14: Inclusion de datos de condominio
'-------------------------------------------------------------------------------------------'
' JJD: 06/02/14: Inclusion de los datos de las Reservas desde la contabilidad (RESCON)
'-------------------------------------------------------------------------------------------'
' RJG: 12/03/14: Ajuste en datos de las Reservas desde la contabilidad (RESCON). Ajustes    '
'                adicionales de presentación de montos (cambios según nueva ley).           '
'-------------------------------------------------------------------------------------------'
' RJG: 21/13/14: Se agregaron parámetros para indicar si se muestran ls locales principales,'
'                adicionales, o todos.                                                      '
'-------------------------------------------------------------------------------------------'
' RJG: 29/04/14: Ajuste en tamaño de fuente para impresión.                                 '
'-------------------------------------------------------------------------------------------'
