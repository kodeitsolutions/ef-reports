'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "Ony_rMargen_gClientes"
'-------------------------------------------------------------------------------------------'
Partial Class Ony_rMargen_gClientes
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
    Dim loAppExcel As Excel.Application


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
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
            Dim lcParametro14Desde As String = cusAplicacion.goReportes.paParametrosIniciales(14)
            Dim lcParametro15Desde As String = cusAplicacion.goReportes.paParametrosIniciales(15)
            Dim lcParametro16Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(16))
            Dim lcParametro16Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(16))
            Dim lcParametro17Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(17))
            Dim lcParametro17Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(17))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim lcCosto As String = ""

            Select Case lcParametro10Desde
                Case "Promedio MP"
                    lcCosto = "Cos_Pro1"
                Case "Ultimo MP"
                    lcCosto = "Cos_Ult1"
                Case "Anterior MP"
                    lcCosto = "Cos_Ant1"
                Case "Promedio MS"
                    lcCosto = "Cos_Pro2"
                Case "Ultimo MS"
                    lcCosto = "Cos_Ult2"
                Case "Anterior MS"
                    lcCosto = "Cos_Ant2"
            End Select

		    Dim llGananciasRespectoAlCosto AS Boolean  = goOpciones.mObtener("GANCOSPRE", "L")

            Dim loComandoSeleccionar As New StringBuilder()


 
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpGanancia(	Cod_Cli		CHAR(10), 			")
            loComandoSeleccionar.AppendLine("							Nom_Cli		CHAR(100),			")
            loComandoSeleccionar.AppendLine("							Base_A		DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Base_B		DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Costo_A		DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Costo_B		DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Can_Fac		INTEGER,	        ")
            loComandoSeleccionar.AppendLine("							Can_Dev		INTEGER)	        ")
            loComandoSeleccionar.AppendLine("															")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpFinal(	Cod_Cli		CHAR(10), 		    ")
            loComandoSeleccionar.AppendLine("							Nom_Cli		CHAR(100),			")
            loComandoSeleccionar.AppendLine("							Base_A		DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Base_B		DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Costo_A		DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Costo_B		DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Ganancia_A	DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Ganancia_B	DECIMAL(28, 10),	")
            loComandoSeleccionar.AppendLine("							Can_Fac		INTEGER,	        ")
            loComandoSeleccionar.AppendLine("							Can_Dev		INTEGER)	        ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
            loComandoSeleccionar.AppendLine("/* Datos de Venta									 										*/")
            loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpGanancia(Cod_Cli, Nom_Cli, Base_A, Base_B, Costo_A, Costo_B, Can_Fac, Can_Dev)")
            loComandoSeleccionar.AppendLine("SELECT			Clientes.Cod_Cli 												        AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli 												        AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("				SUM(  Renglones_Facturas.Mon_Net")
            loComandoSeleccionar.AppendLine("				    *(1-Facturas.por_des1/100+facturas.por_rec1/100) ")
            loComandoSeleccionar.AppendLine("				    *(1+")
            loComandoSeleccionar.AppendLine("				        CASE WHEN Facturas.Mon_Bru>0 ")
            loComandoSeleccionar.AppendLine("				            THEN (Facturas.mon_otr1+Facturas.mon_otr2+Facturas.mon_otr3)/Facturas.Mon_Bru")
            loComandoSeleccionar.AppendLine("				            ELSE 0")
            loComandoSeleccionar.AppendLine("				        END")
            loComandoSeleccionar.AppendLine("				    )) AS Base_A,")
            loComandoSeleccionar.AppendLine("				SUM(COALESCE(Devoluciones.mon_net, 0))                                  AS Base_B,")
            loComandoSeleccionar.AppendLine("				SUM(Renglones_Facturas.Can_Art1*Renglones_Facturas." & lcCosto & ")	    AS Costo_A,")
            loComandoSeleccionar.AppendLine("				SUM(COALESCE(Devoluciones.Can_Art1*Devoluciones." & lcCosto & ", 0))	AS Costo_B,")
            loComandoSeleccionar.AppendLine("				SUM(CASE WHEN Renglones_Facturas.renglon = 1 THEN 1 ELSE 0 END)         AS Can_Fac,")
            loComandoSeleccionar.AppendLine("				SUM(CASE WHEN Devoluciones.renglon = 1 THEN 1 ELSE 0 END)               AS Can_Dev")
            loComandoSeleccionar.AppendLine("FROM			Facturas")
            loComandoSeleccionar.AppendLine(" 		JOIN 	Clientes")
            loComandoSeleccionar.AppendLine(" 			ON	Clientes.Cod_Cli = Facturas.Cod_Cli")
            loComandoSeleccionar.AppendLine(" 		JOIN 	Renglones_Facturas ")
            loComandoSeleccionar.AppendLine(" 			ON	Renglones_Facturas.Documento = Facturas.Documento")
            loComandoSeleccionar.AppendLine(" 	LEFT JOIN 	(	SELECT		Renglones_dClientes.Doc_Ori,")
            loComandoSeleccionar.AppendLine(" 								Renglones_dClientes.Ren_Ori,")
            loComandoSeleccionar.AppendLine(" 								Renglones_dClientes.Can_Art1,")
            loComandoSeleccionar.AppendLine(" 								Renglones_dClientes.Cos_Pro1,")
            loComandoSeleccionar.AppendLine(" 								(   Renglones_dClientes.Mon_Net")
            loComandoSeleccionar.AppendLine(" 								    *(1-Devoluciones_Clientes.por_des1/100+Devoluciones_Clientes.por_rec1/100)")
            loComandoSeleccionar.AppendLine(" 								    *(1+ ")
            loComandoSeleccionar.AppendLine(" 								        CASE WHEN Devoluciones_Clientes.Mon_Bru>0")
            loComandoSeleccionar.AppendLine(" 								        THEN ( Devoluciones_Clientes.mon_otr1")
            loComandoSeleccionar.AppendLine(" 								              +Devoluciones_Clientes.mon_otr2")
            loComandoSeleccionar.AppendLine(" 								              +Devoluciones_Clientes.mon_otr3")
            loComandoSeleccionar.AppendLine(" 								             )/Devoluciones_Clientes.Mon_Bru")
            loComandoSeleccionar.AppendLine(" 								        ELSE 0 END")
            loComandoSeleccionar.AppendLine(" 								)) AS Mon_Net,")
            loComandoSeleccionar.AppendLine(" 								Renglones_dClientes.Renglon")
            loComandoSeleccionar.AppendLine(" 					FROM		Devoluciones_Clientes")
            loComandoSeleccionar.AppendLine(" 						JOIN	Renglones_dClientes ")
            loComandoSeleccionar.AppendLine(" 							ON	Renglones_dClientes.Documento = Devoluciones_Clientes.Documento")
            loComandoSeleccionar.AppendLine(" 					WHERE		Devoluciones_Clientes.Status IN ('Confirmado','Procesado','Afectado')")
            loComandoSeleccionar.AppendLine("                           AND	Renglones_dClientes.Tip_Ori = 'Facturas'")
            loComandoSeleccionar.AppendLine(" 				) AS Devoluciones")
            loComandoSeleccionar.AppendLine(" 			ON	Devoluciones.Doc_Ori = Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine(" 			AND	Devoluciones.Ren_Ori = Renglones_Facturas.Renglon")
            loComandoSeleccionar.AppendLine(" 		JOIN 	Vendedores ")
            loComandoSeleccionar.AppendLine(" 			ON	Vendedores.Cod_Ven = Facturas.Cod_Ven")
            loComandoSeleccionar.AppendLine(" 		JOIN	Articulos ")
            loComandoSeleccionar.AppendLine(" 			ON	Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE		Facturas.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Tip BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Cla BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Ven BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Vendedores.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Status In ( " & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Renglones_Facturas.Cod_Art BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Mon BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Tra BETWEEN " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_For BETWEEN " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Rev BETWEEN " & lcParametro16Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro16Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Suc BETWEEN " & lcParametro17Desde)
            loComandoSeleccionar.AppendLine("    	    AND " & lcParametro17Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Clientes.Cod_Cli, Clientes.Nom_Cli")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
            loComandoSeleccionar.AppendLine("/* Cálculo de ganancia								 										*/ ")
            loComandoSeleccionar.AppendLine("/*------------------------------------------------------------------------------------------*/")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpFinal(Cod_Cli, Nom_Cli, Base_A, Base_B, Costo_A, Costo_B, Ganancia_A, Ganancia_B, Can_Fac, Can_Dev)")
            loComandoSeleccionar.AppendLine("SELECT	Cod_Cli					AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		Nom_Cli					AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		SUM(Base_A)				AS Base_A,")
            loComandoSeleccionar.AppendLine("		SUM(Base_B)				AS Base_B,")
            loComandoSeleccionar.AppendLine("		SUM(Costo_A)			AS Costo_A,")
            loComandoSeleccionar.AppendLine("		SUM(Costo_B)			AS Costo_B,")
            loComandoSeleccionar.AppendLine("		0						AS Ganancia_A,")
            loComandoSeleccionar.AppendLine("		0						AS Ganancia_B,")
            loComandoSeleccionar.AppendLine("		SUM(Can_Fac)            AS Can_Fac,")
            loComandoSeleccionar.AppendLine("		SUM(Can_Dev)            AS Can_Dev")
            loComandoSeleccionar.AppendLine("FROM	#tmpGanancia")
            loComandoSeleccionar.AppendLine("GROUP BY	Cod_Cli, Nom_Cli")
            loComandoSeleccionar.AppendLine("")
            If llGananciasRespectoAlCosto Then 
                loComandoSeleccionar.AppendLine("UPDATE		#tmpFinal")
                loComandoSeleccionar.AppendLine("SET		Ganancia_A = (Base_A -Base_B) - (Costo_A - Costo_B),")
                loComandoSeleccionar.AppendLine("			Ganancia_B = (	CASE	")
                loComandoSeleccionar.AppendLine("								WHEN (Costo_A - Costo_B) <> 0")
                loComandoSeleccionar.AppendLine("								THEN ( (Base_A -Base_B) - (Costo_A - Costo_B))*100 / (Costo_A - Costo_B)")
                loComandoSeleccionar.AppendLine("								ELSE 0")
                loComandoSeleccionar.AppendLine("							END)")
            Else
                loComandoSeleccionar.AppendLine("UPDATE		#tmpFinal")
                loComandoSeleccionar.AppendLine("SET		Ganancia_A = (Base_A -Base_B) - (Costo_A - Costo_B),")
                loComandoSeleccionar.AppendLine("			Ganancia_B = (	CASE	")
                loComandoSeleccionar.AppendLine("								WHEN (Base_A - Base_B) <> 0")
                loComandoSeleccionar.AppendLine("								THEN ( (Base_A -Base_B) - (Costo_A - Costo_B))*100 / (Base_A - Base_B)")
                loComandoSeleccionar.AppendLine("								ELSE 0")
                loComandoSeleccionar.AppendLine("							END)")
            End If
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Cod_Cli					AS Cod_Cli,")
        	loComandoSeleccionar.AppendLine("		Nom_Cli					AS Nom_Cli,")
        	loComandoSeleccionar.AppendLine("		Base_A					AS Base_A,")
        	loComandoSeleccionar.AppendLine("		Base_B					AS Base_B,")
        	loComandoSeleccionar.AppendLine("		Costo_A					AS Costo_A,")
        	loComandoSeleccionar.AppendLine("		Costo_B					AS Costo_B,")
            loComandoSeleccionar.AppendLine("		Ganancia_A				AS Ganancia_A,")
            loComandoSeleccionar.AppendLine("		Ganancia_B				AS Ganancia_B,")
            loComandoSeleccionar.AppendLine("		Can_Fac                 AS Can_Fac,")
            loComandoSeleccionar.AppendLine("		Can_Dev                 AS Can_Dev,")
            If llGananciasRespectoAlCosto Then
			    loComandoSeleccionar.AppendLine("		CAST(1 AS BIT)			AS Ganancia_SobreCosto")
            Else
			    loComandoSeleccionar.AppendLine("		CAST(0 AS BIT)			AS Ganancia_SobreCosto")
            End If
        	loComandoSeleccionar.AppendLine("FROM	#tmpFinal")
            Select Case lcParametro14Desde
                Case "Mayor"
                    loComandoSeleccionar.AppendLine("WHERE Ganancia_B > " & lcParametro15Desde)
                Case "Menor"
                    loComandoSeleccionar.AppendLine("WHERE Ganancia_B < " & lcParametro15Desde)
                Case "Igual"
                    loComandoSeleccionar.AppendLine("WHERE Ganancia_B = " & lcParametro15Desde)
                Case "Todos"
					'No filtra por Ganancia_B
            End Select
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")




            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine(" SELECT COUNT(Case When Ren_Ori = 1 Then 1 end ) As Num_Doc, Doc_Ori, Ren_Ori, Tip_Ori, SUM(Can_Art1) As Can_Art1, SUM(Renglones_dClientes.Mon_Net) As Mon_Net, Renglones_dClientes." & lcCosto)
            'loComandoSeleccionar.AppendLine(" INTO #temDevoluciones")
            'loComandoSeleccionar.AppendLine(" FROM Devoluciones_Clientes")
            'loComandoSeleccionar.AppendLine(" Join Renglones_dClientes ON Renglones_dClientes.Documento = Devoluciones_Clientes.Documento")
            'loComandoSeleccionar.AppendLine(" WHERE Devoluciones_Clientes.Status IN ('Confirmado', 'Afectado', 'Procesado') And Renglones_dClientes.tip_Ori = 'Facturas'")
            'loComandoSeleccionar.AppendLine(" GROUP BY Doc_Ori, Ren_Ori, Tip_Ori, Renglones_dClientes." & lcCosto)
            'loComandoSeleccionar.AppendLine(" ORDER BY Doc_Ori, Ren_Ori, Tip_Ori, Renglones_dClientes." & lcCosto)


            'loComandoSeleccionar.AppendLine(" SELECT")
            'loComandoSeleccionar.AppendLine("             Facturas.Cod_Cli,")
            'loComandoSeleccionar.AppendLine("             COUNT(Distinct Facturas.Documento) AS Can_Fac,")
            'loComandoSeleccionar.AppendLine("             (Select COUNT(devoluciones_clientes.Documento) From devoluciones_clientes Where devoluciones_clientes.Cod_Cli = Facturas.Cod_Cli) AS Can_Dev,")
            'loComandoSeleccionar.AppendLine("             Clientes.Nom_Cli,")
            'loComandoSeleccionar.AppendLine("             SUM(Renglones_Facturas.Mon_Net) AS Base_A,")
            'loComandoSeleccionar.AppendLine("             SUM(CASE")
            'loComandoSeleccionar.AppendLine("					WHEN #temDevoluciones.Tip_Ori = 'Facturas' THEN #temDevoluciones.Mon_Net")
            'loComandoSeleccionar.AppendLine("					ELSE 0")
            'loComandoSeleccionar.AppendLine("			  END) AS Base_B, ")
            'loComandoSeleccionar.AppendLine("             SUM(Renglones_Facturas.Can_Art1 * Renglones_Facturas." & lcCosto & ") AS Costo_A,")
            'loComandoSeleccionar.AppendLine("             SUM(CASE")
            'loComandoSeleccionar.AppendLine("                     WHEN #temDevoluciones.Tip_Ori = 'Facturas' THEN #temDevoluciones." & lcCosto & " * #temDevoluciones.Can_Art1")
            'loComandoSeleccionar.AppendLine("                     ELSE 0")
            'loComandoSeleccionar.AppendLine("             END) AS Costo_B,")
            'loComandoSeleccionar.AppendLine("             SUM(Renglones_Facturas.Mon_Net) - SUM(Renglones_Facturas.Can_Art1 * Renglones_Facturas." & lcCosto & ") AS Ganancia_A,")
            'loComandoSeleccionar.AppendLine("                (((SUM(Renglones_Facturas.Mon_Net) - SUM(Renglones_Facturas.Can_Art1 * Renglones_Facturas." & lcCosto & "))/CASE WHEN SUM(Renglones_Facturas.Mon_Net) = 0 THEN 1 ELSE SUM(Renglones_Facturas.Mon_Net) END)*100)  AS Ganancia_B")
            'loComandoSeleccionar.AppendLine(" INTO           #Temp")
            'loComandoSeleccionar.AppendLine(" FROM           Facturas")
            'loComandoSeleccionar.AppendLine(" JOIN Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            'loComandoSeleccionar.AppendLine(" JOIN Vendedores ON Vendedores.Cod_Ven = Facturas.Cod_Ven")
            'loComandoSeleccionar.AppendLine(" JOIN Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            'loComandoSeleccionar.AppendLine(" LEFT JOIN #temDevoluciones ON ")
            'loComandoSeleccionar.AppendLine(" 			#temDevoluciones.Doc_Ori = Facturas.Documento ")
            'loComandoSeleccionar.AppendLine(" 			AND #temDevoluciones.Doc_Ori = Renglones_Facturas.Documento ")
            'loComandoSeleccionar.AppendLine(" 			AND #temDevoluciones.Ren_Ori = Renglones_Facturas.Renglon ")
            'loComandoSeleccionar.AppendLine(" JOIN Articulos ON Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            'loComandoSeleccionar.AppendLine(" WHERE")



            'loComandoSeleccionar.AppendLine(" 			Facturas.Documento BETWEEN " & lcParametro0Desde)
            'loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            'loComandoSeleccionar.AppendLine(" 			AND Facturas.Fec_Ini BETWEEN " & lcParametro1Desde)
            'loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            'loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Cli BETWEEN " & lcParametro2Desde)
            'loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            'loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Tip BETWEEN " & lcParametro3Desde)
            'loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            'loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Cla BETWEEN " & lcParametro4Desde)
            'loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            'loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Ven BETWEEN " & lcParametro5Desde)
            'loComandoSeleccionar.AppendLine("    	    AND " & lcParametro5Hasta)
            'loComandoSeleccionar.AppendLine(" 			AND Vendedores.Cod_Tip BETWEEN " & lcParametro6Desde)
            'loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            'loComandoSeleccionar.AppendLine(" 			AND Facturas.Status In ( " & lcParametro7Desde & ")")
            'loComandoSeleccionar.AppendLine("           AND Renglones_Facturas.Cod_Art BETWEEN " & lcParametro8Desde)
            'loComandoSeleccionar.AppendLine("    	    AND " & lcParametro8Hasta)
            'loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep BETWEEN " & lcParametro9Desde)
            'loComandoSeleccionar.AppendLine("    	    AND " & lcParametro9Hasta)
            'loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Mon BETWEEN " & lcParametro11Desde)
            'loComandoSeleccionar.AppendLine("    	    AND " & lcParametro11Hasta)
            'loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Tra BETWEEN " & lcParametro12Desde)
            'loComandoSeleccionar.AppendLine("    	    AND " & lcParametro12Hasta)
            'loComandoSeleccionar.AppendLine("           AND Facturas.Cod_For BETWEEN " & lcParametro13Desde)
            'loComandoSeleccionar.AppendLine("    	    AND " & lcParametro13Hasta)
            'loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Rev BETWEEN " & lcParametro16Desde)
            'loComandoSeleccionar.AppendLine("    	    AND " & lcParametro16Hasta)
            'loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Suc BETWEEN " & lcParametro17Desde)
            'loComandoSeleccionar.AppendLine("    	    AND " & lcParametro17Hasta)

            'loComandoSeleccionar.AppendLine(" GROUP BY    Facturas.Cod_Cli, Clientes.Nom_Cli")

            'loComandoSeleccionar.AppendLine(" ORDER BY     Facturas.Cod_Cli, Clientes.Nom_Cli")


            'Select Case lcParametro14Desde
            '    Case "Mayor"
            '        loComandoSeleccionar.AppendLine("SELECT")
            '        loComandoSeleccionar.AppendLine("            Cod_Cli,")
            '        loComandoSeleccionar.AppendLine("            Nom_Cli,")
            '        loComandoSeleccionar.AppendLine("            SUM(Can_Fac) AS Can_Fac,")
            '        loComandoSeleccionar.AppendLine("            SUM(Can_Dev) AS Can_Dev,")
            '        loComandoSeleccionar.AppendLine("            SUM(Base_A) AS Base_A, ")
            '        loComandoSeleccionar.AppendLine("            SUM(Base_B) AS Base_B,")
            '        loComandoSeleccionar.AppendLine("            SUM(Costo_a) AS Costo_A,")
            '        loComandoSeleccionar.AppendLine("            SUM(Costo_B) AS Costo_B,")
            '        loComandoSeleccionar.AppendLine("            SUM(Ganancia_A) AS Ganancia_A,")
            '        loComandoSeleccionar.AppendLine("            AVG(Ganancia_B) AS Ganancia_B")
            '        loComandoSeleccionar.AppendLine(" FROM #Temp ")
            '        loComandoSeleccionar.AppendLine(" WHERE Ganancia_B > " & lcParametro15Desde)
            '        loComandoSeleccionar.AppendLine("  GROUP BY Cod_Cli, Nom_Cli")
            '    Case "Menor"
            '        loComandoSeleccionar.AppendLine("SELECT")
            '        loComandoSeleccionar.AppendLine("            Cod_Cli,")
            '        loComandoSeleccionar.AppendLine("            Nom_Cli,")
            '        loComandoSeleccionar.AppendLine("            SUM(Can_Fac) AS Can_Fac,")
            '        loComandoSeleccionar.AppendLine("            SUM(Can_Dev) AS Can_Dev,")
            '        loComandoSeleccionar.AppendLine("            SUM(Base_A) AS Base_A, ")
            '        loComandoSeleccionar.AppendLine("            SUM(Base_B) AS Base_B,")
            '        loComandoSeleccionar.AppendLine("            SUM(Costo_a) AS Costo_A,")
            '        loComandoSeleccionar.AppendLine("            SUM(Costo_B) AS Costo_B,")
            '        loComandoSeleccionar.AppendLine("            SUM(Ganancia_A) AS Ganancia_A,")
            '        loComandoSeleccionar.AppendLine("            AVG(Ganancia_B) AS Ganancia_B")
            '        loComandoSeleccionar.AppendLine(" FROM #Temp ")
            '        loComandoSeleccionar.AppendLine(" WHERE Ganancia_B < " & lcParametro15Desde)
            '        loComandoSeleccionar.AppendLine("  GROUP BY Cod_Cli, Nom_Cli")
            '    Case "Igual"
            '        loComandoSeleccionar.AppendLine("SELECT")
            '        loComandoSeleccionar.AppendLine("            Cod_Cli,")
            '        loComandoSeleccionar.AppendLine("            Nom_Cli,")
            '        loComandoSeleccionar.AppendLine("            SUM(Can_Fac) AS Can_Fac,")
            '        loComandoSeleccionar.AppendLine("            SUM(Can_Dev) AS Can_Dev,")
            '        loComandoSeleccionar.AppendLine("            SUM(Base_A) AS Base_A, ")
            '        loComandoSeleccionar.AppendLine("            SUM(Base_B) AS Base_B,")
            '        loComandoSeleccionar.AppendLine("            SUM(Costo_a) AS Costo_A,")
            '        loComandoSeleccionar.AppendLine("            SUM(Costo_B) AS Costo_B,")
            '        loComandoSeleccionar.AppendLine("            SUM(Ganancia_A) AS Ganancia_A,")
            '        loComandoSeleccionar.AppendLine("            AVG(Ganancia_B) AS Ganancia_B")
            '        loComandoSeleccionar.AppendLine(" FROM #Temp ")
            '        loComandoSeleccionar.AppendLine(" WHERE Ganancia_B = " & lcParametro15Desde)
            '        loComandoSeleccionar.AppendLine("  GROUP BY Cod_Cli, Nom_Cli")
            '    Case "Todos"
            '        loComandoSeleccionar.AppendLine("SELECT")
            '        loComandoSeleccionar.AppendLine("            Cod_Cli,")
            '        loComandoSeleccionar.AppendLine("            Nom_Cli,")
            '        loComandoSeleccionar.AppendLine("            SUM(Can_Fac) AS Can_Fac,")
            '        loComandoSeleccionar.AppendLine("            SUM(Can_Dev) AS Can_Dev,")
            '        loComandoSeleccionar.AppendLine("            SUM(Base_A) AS Base_A, ")
            '        loComandoSeleccionar.AppendLine("            SUM(Base_B) AS Base_B,")
            '        loComandoSeleccionar.AppendLine("            SUM(Costo_a) AS Costo_A,")
            '        loComandoSeleccionar.AppendLine("            SUM(Costo_B) AS Costo_B,")
            '        loComandoSeleccionar.AppendLine("            SUM(Ganancia_A) AS Ganancia_A,")
            '        loComandoSeleccionar.AppendLine("            AVG(Ganancia_B) AS Ganancia_B")
            '        loComandoSeleccionar.AppendLine(" FROM #Temp ")
            '        loComandoSeleccionar.AppendLine("  GROUP BY Cod_Cli, Nom_Cli")
            'End Select

            'loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("Ony_rMargen_gClientes", laDatosReporte)

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
            Me.crvOny_rMargen_gClientes.ReportSource = loObjetoReporte


            'Selección de opcion por excel (Microsoft Excel - xls)
            If (Me.Request.QueryString("salida").ToLower = "xls") Then
                ' Ruta donde se creara temporalmente el archivo
                Dim lcFileName As String = Server.MapPath("~\Administrativo\Temporales\Ony_rMargen_gClientes_" & Guid.NewGuid().ToString("N") & ".xls")
                ' Se exporta para crear el archivo temporal
                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, lcFileName)

                ' Se modifica el contenido del archivo
                Me.modificar_excel(lcFileName, laDatosReporte)

                ' Se coloca en la respuesta para decargar
                Me.Response.Clear()
                Me.Response.Buffer = True
                Me.Response.AppendHeader("content-disposition", "attachment; filename=Ony_rMargen_gClientes.xls")
                Me.Response.ContentType = "application/excel"
                Me.Response.WriteFile(lcFileName, True)
                Me.Response.Write(Space(30))
                Me.Response.Flush()
                Me.Response.Close()

                Me.Response.End()

            End If

        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                 "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                  vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                  "auto", _
                  "auto")

        End Try

    End Sub

    ''' <summary>
    ''' Rutina que modifica y formatea el contenido del archivo excel a descargar.
    ''' </summary>
    ''' <param name="loFileName">Ruta del archivo a modificar.</param>
    ''' <param name="loDatosReporte">Datos de la consulta a la base de datos.</param>
    ''' <remarks></remarks>
    Private Sub modificar_excel(ByVal loFileName As String, ByVal loDatosReporte As DataSet)

		Dim llGananciasRespectoAlCosto AS Boolean = goOpciones.mObtener("GANCOSPRE", "L")
		

        Try
            ' Se inicializa el objeto a la aplicacion excel
            loAppExcel = New Excel.Application()
            loAppExcel.Visible = False
            loAppExcel.DisplayAlerts = False

            ' Se carga el archivo excel a modificar
            Dim lcLibrosExcel As Excel.Workbooks = loAppExcel.Workbooks
            Dim lcLibroExcel As Excel.Workbook = lcLibrosExcel.Open(loFileName)

            ' Se activa la primera hoja del libro donde se almacenara toda la informacion
            Dim lcHojaExcel As Excel.Worksheet
            lcHojaExcel = lcLibroExcel.Worksheets(1)
            lcHojaExcel.Activate()

            ' Se selecciona toda la hoja para blanquera todo
            '   - El número total de columnas disponibles en Excel
            '       Viejo límite: 256 (2^8)         (Excel 2003 o inferor)
            '       Nuevo límite: 16.384 (2^14)     (Excel 2007)
            '   - El número total de filas disponibles en Excel
            '       Viejo límite: 65.536 (2^16)         (Excel 2003 o inferior)
            '       Nuevo límite: el 1.048.576 (2^20)   (Excel 2007)
            Dim lcRango As Excel.Range = lcHojaExcel.Range("A1:IV65536")
            lcRango.Select()
            lcRango.Clear()
            lcRango.Font.Size = 9
            lcRango.Font.Name = "Tahoma"

            ' Nombre de la empresa
            lcHojaExcel.Cells(1, 1).Value = cusAplicacion.goEmpresa.pcNombre.ToUpper
            ' Nombre del modulo
            lcHojaExcel.Cells(2, 1).Value = "Cuentas x Cobrar"
            ' Titulo del reporte
            lcRango = lcHojaExcel.Range("A3:L3")
            lcRango.Select()
            lcRango.Font.ColorIndex = 25
            lcRango.Interior.ColorIndex = 34
            lcRango.MergeCells = True
            lcRango.Value = Me.pcNombreReporte
            lcRango.Font.Size = 14
            lcRango.Font.Bold = True
            lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            lcRango.Rows.AutoFit()

            ' Parametros del reporte
            lcRango = lcHojaExcel.Range("A4:L4")
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.Value = cusAplicacion.goReportes.mObtenerParametros(cusAplicacion.goReportes.paNombresParametros, cusAplicacion.goReportes.paParametrosIniciales, cusAplicacion.goReportes.paParametrosFinales)
            lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
            lcRango.Rows.AutoFit()

            lcHojaExcel.Cells(6, 3).Value = "Monto Base"
            lcRango = lcHojaExcel.Range("C6:F6")
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.Font.ColorIndex = 25
            lcRango.Interior.ColorIndex = 34

            lcHojaExcel.Cells(6, 8).Value = "Costo"
            lcRango = lcHojaExcel.Range("H6:I6")
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.Font.ColorIndex = 25
            lcRango.Interior.ColorIndex = 34

            lcHojaExcel.Cells(6, 11).Value = "Ganancia"
            lcRango = lcHojaExcel.Range("K6:L6")
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.Font.ColorIndex = 25
            lcRango.Interior.ColorIndex = 34

            lcRango = lcHojaExcel.Range("A6:L6")
            lcRango.Select()
            lcRango.Font.Bold = True
            lcRango.EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            lcRango.EntireRow.WrapText = True
            


            lcHojaExcel.Cells(7, 1).Value = "Código"
            lcHojaExcel.Cells(7, 2).Value = "Nombre"
            lcHojaExcel.Cells(7, 3).Value = "Monto Total"
            lcHojaExcel.Cells(7, 4).Value = "# Fac."
            lcHojaExcel.Cells(7, 5).Value = "Devoluciones"
            lcHojaExcel.Cells(7, 6).Value = "# Dev."
            lcHojaExcel.Cells(7, 7).Value = "Monto Total - Devoluciones"
            lcHojaExcel.Cells(7, 8).Value = "Ventas"
            lcHojaExcel.Cells(7, 9).Value = "Devoluciones"
            lcHojaExcel.Cells(7, 10).Value = "Ventas - Devoluciones"
            lcHojaExcel.Cells(7, 11).Value = "Monto"
            lcHojaExcel.Cells(7, 12).Value = "%"
            
            ' Se le da formato a las celdas del membrete de la tabla

            lcRango = lcHojaExcel.Range("A7:L7")
            lcRango.Select()
            lcRango.Font.Bold = True
            lcRango.EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            lcRango.EntireRow.WrapText = True
            lcRango.Font.ColorIndex = 34
            lcRango.Interior.ColorIndex = 25

            ' Formato a las columnas de la tabla de informacion del reporte

            lcRango = lcHojaExcel.Range("A1:B1")
            lcRango.EntireColumn.NumberFormat = "@"

            lcRango = lcHojaExcel.Range("C1")
            lcRango.EntireColumn.NumberFormat = "###,###,##0.00"

            lcRango = lcHojaExcel.Range("D1")
            lcRango.EntireColumn.NumberFormat = "###,###,##0"
            lcRango = lcHojaExcel.Range("E1")
            lcRango.EntireColumn.NumberFormat = "###,###,##0.00"
            lcRango = lcHojaExcel.Range("F1")
            lcRango.EntireColumn.NumberFormat = "###,###,##0"
            lcRango = lcHojaExcel.Range("G1:L1")
            lcRango.EntireColumn.NumberFormat = "###,###,##0.00"

            ' Formato a las celdas de la fecha y hora de creacion
            ' Fecha y hora de creacion
            Dim lcFechaCreacion As DateTime = DateTime.Now()
            lcHojaExcel.Cells(1, 12).NumberFormat = "@"
            lcHojaExcel.Cells(1, 12).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            lcHojaExcel.Cells(1, 12).Value = lcFechaCreacion.ToString("dd/MM/yyyy")
            lcHojaExcel.Cells(2, 12).NumberFormat = "@"
            lcHojaExcel.Cells(2, 12).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            lcHojaExcel.Cells(2, 12).Value = lcFechaCreacion.ToString("hh:mm:ss tt")

            ' Recorrido de los datos de la consulta a la base de datos para introducir en la tabla del reporte

            ' Numero de la fila de los datos obtenidos de la consulta a la base de datos
            Dim lcFila As Integer
            ' Total de filas de los datos obtenidos de la consulta a la base de datos
            Dim lcTotalFilas As Integer = loDatosReporte.Tables(0).Rows.Count - 1
            ' Datos de la fila de la consulta a la base de datos
            Dim lcDatosFila As DataRow
            '' Control de agrupamiento 1 - (Clase de Articulo)(Cod_Cla)
            'Dim grupo1 As String = ""
            '' Control de agrupamiento 2 - (Codigo del Tipo de Documento)(Cod_Tip)
            'Dim grupo2 As String = ""
            ' Numero de la fila en el documento excel
            Dim lcNumFila As Integer = 8
            ' Numero de la fila en el documento excel, inicial para la sumatoria del total para el grupo de control 2
            Dim lcFilaIni As Integer = 8
            ' Numero de la fila en el documento excel, final para la sumatoria del total para el grupo de control 2
            Dim lcFilaFin As Integer = 0
            ' Construccion de la formula de total para el grupo de control 1
            Dim lcTotalDocumento As String = "= 0"

            ' Recorriendo las filas de los datos de la consulta a la base de datos
            For lcFila = 0 To lcTotalFilas
                ' Se extrae los datos de la fila
                lcDatosFila = loDatosReporte.Tables(0).Rows(lcFila)

                ' Se agrega la informacion en la tabla de los valores detallados

                lcHojaExcel.Cells(lcNumFila, 1).Value = lcDatosFila("Cod_Cli")
                lcHojaExcel.Cells(lcNumFila, 2).Value = lcDatosFila("Nom_Cli")
                lcHojaExcel.Cells(lcNumFila, 3).Value = lcDatosFila("Base_A")
                lcHojaExcel.Cells(lcNumFila, 4).Value = lcDatosFila("Can_Fac")
                lcHojaExcel.Cells(lcNumFila, 5).Value = lcDatosFila("Base_B")
                lcHojaExcel.Cells(lcNumFila, 6).Value = lcDatosFila("Can_Dev")
                lcHojaExcel.Cells(lcNumFila, 7).Value = lcDatosFila("Base_A") - lcDatosFila("Base_B")
                lcHojaExcel.Cells(lcNumFila, 8).Value = lcDatosFila("Costo_A")
                lcHojaExcel.Cells(lcNumFila, 9).Value = lcDatosFila("Costo_B")
                lcHojaExcel.Cells(lcNumFila, 10).Value = lcDatosFila("Costo_A") - lcDatosFila("Costo_B")
                lcHojaExcel.Cells(lcNumFila, 11).Value = lcDatosFila("Ganancia_A")
                lcHojaExcel.Cells(lcNumFila, 12).Value = lcDatosFila("Ganancia_B")
               
                lcNumFila = lcNumFila + 1

            Next lcFila

            ' Se almacena el numero de la fila final del grupo de datos
            lcFilaFin = lcNumFila - 1
            ' Se coloca la etiqueta de total 
            lcRango = lcHojaExcel.Range("B" & CStr(lcNumFila) & ":B" & CStr(lcNumFila))
            lcRango.Select()
            lcRango.MergeCells = True
            lcRango.EntireRow.Font.Bold = True
            lcRango.Value = "Totales:"

            ' Se coloca la formula de total de todas las columnas


            lcHojaExcel.Cells(lcNumFila, 3).Formula = "=SUM(C" & CStr(lcFilaIni) & ":C" & CStr(lcNumFila - 1) & ")"
            lcHojaExcel.Cells(lcNumFila, 4).Formula = "=SUM(D" & CStr(lcFilaIni) & ":D" & CStr(lcNumFila - 1) & ")"
            lcHojaExcel.Cells(lcNumFila, 5).Formula = "=SUM(E" & CStr(lcFilaIni) & ":E" & CStr(lcNumFila - 1) & ")"
            lcHojaExcel.Cells(lcNumFila, 6).Formula = "=SUM(F" & CStr(lcFilaIni) & ":F" & CStr(lcNumFila - 1) & ")"
            lcHojaExcel.Cells(lcNumFila, 7).Formula = "=SUM(G" & CStr(lcFilaIni) & ":G" & CStr(lcNumFila - 1) & ")"
            lcHojaExcel.Cells(lcNumFila, 8).Formula = "=SUM(H" & CStr(lcFilaIni) & ":H" & CStr(lcNumFila - 1) & ")"
            lcHojaExcel.Cells(lcNumFila, 9).Formula = "=SUM(I" & CStr(lcFilaIni) & ":I" & CStr(lcNumFila - 1) & ")"
            lcHojaExcel.Cells(lcNumFila, 10).Formula = "=SUM(J" & CStr(lcFilaIni) & ":J" & CStr(lcNumFila - 1) & ")"
            lcHojaExcel.Cells(lcNumFila, 11).Formula = "=SUM(K" & CStr(lcFilaIni) & ":K" & CStr(lcNumFila - 1) & ")"
            If llGananciasRespectoAlCosto Then 
    			lcHojaExcel.Cells(lcNumFila, 12).Formula = "=IF(J" & (lcNumFila) & ">0, K" & (lcNumFila) & "*100/J" & (lcNumFila) & ", 100)"
            Else
    			lcHojaExcel.Cells(lcNumFila, 12).Formula = "=IF(G" & (lcNumFila) & ">0, K" & (lcNumFila) & "*100/G" & (lcNumFila) & ", 100)"
            End If


            lcNumFila = lcNumFila + 1

            ' Ajustamos el tamaño de las columnas
            lcRango = lcHojaExcel.Range("A1:A" & CStr(lcNumFila))
            lcRango.Select()
            lcRango.ColumnWidth = 11
            lcRango = lcHojaExcel.Range("B1:B" & CStr(lcNumFila))
            lcRango.Select()
            lcRango.ColumnWidth = 35
            lcRango = lcHojaExcel.Range("C1:L" & CStr(lcNumFila))
            lcRango.Select()
            lcRango.ColumnWidth = 12
            
            'lcRango.WrapText = True
            'lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            'lcRango.Rows.AutoFit()


            ' Seleccionamos la primera celda del libro
            lcRango = lcHojaExcel.Range("A1")
            lcRango.Select()

            ' Cerramos y liberamos recursos
            mLiberar(lcRango)
            mLiberar(lcHojaExcel)
            'Guardamos los cambios del libro activo
            lcLibroExcel.Close(True, loFileName)

            mLiberar(lcLibroExcel)
            loAppExcel.Application.Quit()
            mLiberar(loAppExcel)

        Catch loExcepcion As Exception
            Me.mEscribirConsulta(loExcepcion.Message)
            Me.Response.Flush()
            Me.Response.Close()

            Me.Response.End()

        Finally
            ' Se forza el cierre del proceso excel
            'Dim Lista_Procesos() As Diagnostics.Process
            'Dim p As Diagnostics.Process
            'Lista_Procesos = Diagnostics.Process.GetProcessesByName("EXCEL")
            'For Each p In Lista_Procesos
            '    Try
            '        p.Kill()
            '    Catch
            '    End Try
            'Next
            GC.Collect()
        End Try

    End Sub


    ''' <summary>
    ''' Cierre y liberacion de recursos de los objetos de la libreria Excel
    ''' </summary>
    ''' <param name="objeto"></param>
    ''' <remarks></remarks>
    Private Sub mLiberar(ByVal objeto As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto)
            objeto = Nothing
        Catch ex As Exception
            objeto = Nothing
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
' CMS: 30/08/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
' RJG: 16/01/14: Se agregó la opción para el cálculo de ganancias con respecto al precio o  '
'                costo. Se ajustó el SELECT para considerar los Descuentos, Recargos y Otros. 
'-------------------------------------------------------------------------------------------'
