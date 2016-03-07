'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grMargen_gEmpresa"
'-------------------------------------------------------------------------------------------'
Partial Class grMargen_gEmpresa
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load



        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1HastaAux As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
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
            Dim lcCosto As String

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

            Dim loComandoSeleccionar As New StringBuilder()

            'loComandoSeleccionar.AppendLine(" SELECT")
            'loComandoSeleccionar.AppendLine("     Facturas.Fec_Ini,")
            'loComandoSeleccionar.AppendLine("     SUM(Renglones_Facturas.Mon_Net) AS Ventas,")
            'loComandoSeleccionar.AppendLine("     SUM(CASE")
            'loComandoSeleccionar.AppendLine("             WHEN Renglones_dClientes.Tip_Ori = 'Facturas' THEN Renglones_dClientes.Mon_Net")
            'loComandoSeleccionar.AppendLine("             ELSE 0")
            'loComandoSeleccionar.AppendLine("     END) AS Base_B, ")
            'loComandoSeleccionar.AppendLine("     SUM(Renglones_Facturas.Can_Art1 * Renglones_Facturas." & lcCosto & ") AS Costos,")
            'loComandoSeleccionar.AppendLine("     SUM(CASE")
            'loComandoSeleccionar.AppendLine("             WHEN Renglones_dClientes.Tip_Ori = 'Facturas' THEN Renglones_dClientes." & lcCosto & " * Renglones_dClientes.Can_Art1")
            'loComandoSeleccionar.AppendLine("             ELSE 0")
            'loComandoSeleccionar.AppendLine("     END) AS Costo_B,")
            'loComandoSeleccionar.AppendLine("     SUM(Renglones_Facturas.Mon_Net) - SUM(Renglones_Facturas.Can_Art1 * Renglones_Facturas." & lcCosto & ") AS Ganancias,")
            ''loComandoSeleccionar.AppendLine("     (((SUM(Renglones_Facturas.Mon_Net) - SUM(Renglones_Facturas.Can_Art1 * Renglones_Facturas." & lcCosto & "))/SUM(Renglones_Facturas.Mon_Net))*100) AS Ganancia_B")
            'loComandoSeleccionar.AppendLine("                (((SUM(Renglones_Facturas.Mon_Net) - SUM(Renglones_Facturas.Can_Art1 * Renglones_Facturas." & lcCosto & "))/CASE WHEN SUM(Renglones_Facturas.Mon_Net) = 0 THEN 1 ELSE SUM(Renglones_Facturas.Mon_Net) END)*100)  AS Ganancia_B")
            'loComandoSeleccionar.AppendLine(" INTO           #Temp")
            'loComandoSeleccionar.AppendLine(" FROM           Facturas")
            'loComandoSeleccionar.AppendLine(" JOIN Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            'loComandoSeleccionar.AppendLine(" JOIN Vendedores ON Vendedores.Cod_Ven = Facturas.Cod_Ven")
            'loComandoSeleccionar.AppendLine(" JOIN Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            'loComandoSeleccionar.AppendLine(" JOIN Articulos ON Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            'loComandoSeleccionar.AppendLine(" LEFT JOIN Renglones_dClientes ON Renglones_dClientes.Doc_Ori = Facturas.Documento AND Articulos.Cod_Art = Renglones_dClientes.Cod_Art")
            'loComandoSeleccionar.AppendLine(" LEFT JOIN Devoluciones_Clientes ON Devoluciones_Clientes.Documento = Renglones_dClientes.Documento")
            'loComandoSeleccionar.AppendLine(" WHERE")
            
            loComandoSeleccionar.AppendLine("SELECT 	Doc_Ori, Ren_Ori, Tip_Ori, SUM(Can_Art1) As Can_Art1, SUM(Renglones_dClientes.Mon_Net) As Mon_Net, Renglones_dClientes." & lcCosto)
			loComandoSeleccionar.AppendLine("INTO 		#temDevoluciones")
			loComandoSeleccionar.AppendLine("FROM 		Devoluciones_Clientes")
			loComandoSeleccionar.AppendLine("	JOIN	Renglones_dClientes ON Renglones_dClientes.Documento = Devoluciones_Clientes.Documento")
			loComandoSeleccionar.AppendLine("WHERE		Devoluciones_Clientes.Status IN ('Confirmado', 'Afectado', 'Procesado') AND Renglones_dClientes.tip_Ori = 'Facturas'")
			loComandoSeleccionar.AppendLine("GROUP BY 	Doc_Ori, Ren_Ori, Tip_Ori, Renglones_dClientes." & lcCosto)
			loComandoSeleccionar.AppendLine("ORDER BY 	Doc_Ori, Ren_Ori, Tip_Ori, Renglones_dClientes." & lcCosto)
            
            loComandoSeleccionar.AppendLine("SELECT	Facturas.Fec_Ini,")
            loComandoSeleccionar.AppendLine("     	SUM(Renglones_Facturas.Mon_Net) AS Ventas,")
            loComandoSeleccionar.AppendLine("   		SUM(CASE")
            loComandoSeleccionar.AppendLine("				WHEN #temDevoluciones.Tip_Ori = 'Facturas' THEN #temDevoluciones.Mon_Net")
            loComandoSeleccionar.AppendLine("				ELSE 0")
            loComandoSeleccionar.AppendLine("			END) AS Base_B, ")
            loComandoSeleccionar.AppendLine("     	SUM(Renglones_Facturas.Can_Art1 * Renglones_Facturas." & lcCosto & ") AS Costos,")
            loComandoSeleccionar.AppendLine("     	SUM(CASE")
            loComandoSeleccionar.AppendLine("     	        WHEN #temDevoluciones.Tip_Ori = 'Facturas' THEN #temDevoluciones." & lcCosto & " * #temDevoluciones.Can_Art1")
            loComandoSeleccionar.AppendLine("     	        ELSE 0")
            loComandoSeleccionar.AppendLine("     	END) AS Costo_B,")
            loComandoSeleccionar.AppendLine("     	SUM(Renglones_Facturas.Mon_Net) - SUM(Renglones_Facturas.Can_Art1 * Renglones_Facturas." & lcCosto & ") AS Ganancias,")
            loComandoSeleccionar.AppendLine("     	           (((SUM(Renglones_Facturas.Mon_Net) - SUM(Renglones_Facturas.Can_Art1 * Renglones_Facturas." & lcCosto & "))/CASE WHEN SUM(Renglones_Facturas.Mon_Net) = 0 THEN 1 ELSE SUM(Renglones_Facturas.Mon_Net) END)*100)  AS Ganancia_B")
            loComandoSeleccionar.AppendLine("INTO   #Temp")
            loComandoSeleccionar.AppendLine("FROM   Facturas")
            loComandoSeleccionar.AppendLine(" 	JOIN Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            loComandoSeleccionar.AppendLine(" 	JOIN Vendedores ON Vendedores.Cod_Ven = Facturas.Cod_Ven")
            loComandoSeleccionar.AppendLine(" 	JOIN Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
			loComandoSeleccionar.AppendLine("	LEFT JOIN #temDevoluciones ")
			loComandoSeleccionar.AppendLine(" 		ON	#temDevoluciones.Doc_Ori = Facturas.Documento ")
			loComandoSeleccionar.AppendLine(" 		AND #temDevoluciones.Doc_Ori = Renglones_Facturas.Documento ")
			loComandoSeleccionar.AppendLine(" 		AND #temDevoluciones.Ren_Ori = Renglones_Facturas.Renglon ")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE		Facturas.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Fec_Ini BETWEEN DATEADD (mm, -30, " & lcParametro1HastaAux & " )")
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Cli BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Tip BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Cla BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Ven BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("    			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Vendedores.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Status In ( " & lcParametro7Desde & ")")
            loComandoSeleccionar.AppendLine("           AND Renglones_Facturas.Cod_Art BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    			AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("    			AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Mon BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine("    			AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Tra BETWEEN " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine("    			AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_For BETWEEN " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("    			AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Rev BETWEEN " & lcParametro16Desde)
            loComandoSeleccionar.AppendLine("    			AND " & lcParametro16Hasta)
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Suc BETWEEN " & lcParametro17Desde)
            loComandoSeleccionar.AppendLine("    			AND " & lcParametro17Hasta)

            loComandoSeleccionar.AppendLine("GROUP BY		Facturas.Fec_Ini")

            loComandoSeleccionar.AppendLine("ORDER BY		Facturas.Fec_Ini")

           
            Select Case lcParametro14Desde
                Case "Mayor"
                    loComandoSeleccionar.AppendLine("SELECT		DATEPART(MONTH, DATEADD (MONTH , -30, " & lcParametro1Hasta & " )) AS Mes_Inicio,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR, DATEADD (MONTH , -30, " & lcParametro1Hasta & " )) AS Año_Inicio,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(MONTH," & lcParametro1Hasta & ") AS Mes_Fin,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR," & lcParametro1Hasta & ") AS Año_Fin,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR,Fec_Ini) AS Año,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(MONTH,Fec_Ini) AS Mes,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Ventas) AS Ventas, ")
                    loComandoSeleccionar.AppendLine(" 			SUM(Base_B) AS Base_B,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Costos) AS Costos,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Costo_B) AS Costo_B,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Ganancias) AS Ganancias,")
                    loComandoSeleccionar.AppendLine(" 			AVG(Ganancia_B) AS Ganancia_B")
                    loComandoSeleccionar.AppendLine("FROM		#Temp ")
                    loComandoSeleccionar.AppendLine("WHERE		Ganancia_B > " & lcParametro15Desde)
                    loComandoSeleccionar.AppendLine("GROUP BY	DATEPART(YEAR,Fec_Ini), DATEPART(MONTH,Fec_Ini)")
                Case "Menor"
                    loComandoSeleccionar.AppendLine("SELECT		DATEPART(MONTH, DATEADD (MONTH , -30, " & lcParametro1Hasta & " ))	AS Mes_Inicio,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR, DATEADD (MONTH , -30, " & lcParametro1Hasta & " ))	AS Año_Inicio,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(MONTH," & lcParametro1Hasta & ")	AS Mes_Fin,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR," & lcParametro1Hasta & ")	AS Año_Fin,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR,Fec_Ini)						AS Año,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(MONTH,Fec_Ini)						AS Mes,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Ventas) 								AS Ventas, ")
                    loComandoSeleccionar.AppendLine(" 			SUM(Base_B) 								AS Base_B,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Costos) 								AS Costos,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Costo_B)								AS Costo_B,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Ganancias)								AS Ganancias,")
                    loComandoSeleccionar.AppendLine(" 			AVG(Ganancia_B)								AS Ganancia_B")
                    loComandoSeleccionar.AppendLine("FROM		#Temp ")
                    loComandoSeleccionar.AppendLine("WHERE		Ganancia_B < " & lcParametro15Desde)
                    loComandoSeleccionar.AppendLine("GROUP BY	DATEPART(YEAR,Fec_Ini), DATEPART(MONTH,Fec_Ini)")
                Case "Igual"
                    loComandoSeleccionar.AppendLine("SELECT		DATEPART(MONTH, DATEADD (MONTH , -30, " & lcParametro1Hasta & " )) AS Mes_Inicio,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR, DATEADD (MONTH , -30, " & lcParametro1Hasta & " )) AS Año_Inicio,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(MONTH," & lcParametro1Hasta & ")	AS Mes_Fin,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR," & lcParametro1Hasta & ")	AS Año_Fin,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR,Fec_Ini)						AS Año,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(MONTH,Fec_Ini) 					AS Mes,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Ventas) 								AS Ventas, ")
                    loComandoSeleccionar.AppendLine(" 			SUM(Base_B) 								AS Base_B,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Costos) 								AS Costos,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Costo_B)								AS Costo_B,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Ganancias)								AS Ganancias,")
                    loComandoSeleccionar.AppendLine(" 			AVG(Ganancia_B) 							AS Ganancia_B")
                    loComandoSeleccionar.AppendLine("FROM		#Temp ")
                    loComandoSeleccionar.AppendLine("WHERE		Ganancia_B = " & lcParametro15Desde)
                    loComandoSeleccionar.AppendLine("GROUP BY	DATEPART(YEAR,Fec_Ini), DATEPART(MONTH,Fec_Ini)")
                Case "Todos"
                    loComandoSeleccionar.AppendLine("SELECT		DATEPART(MONTH, DATEADD (MONTH , -30, " & lcParametro1Hasta & " )) AS Mes_Inicio,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR, DATEADD (MONTH , -30, " & lcParametro1Hasta & " )) AS Año_Inicio,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(MONTH," & lcParametro1Hasta & ")	AS Mes_Fin,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR," & lcParametro1Hasta & ")	AS Año_Fin,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(YEAR,Fec_Ini)						AS Año,")
                    loComandoSeleccionar.AppendLine(" 			DATEPART(MONTH,Fec_Ini)						AS Mes,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Ventas) 								AS Ventas, ")
                    loComandoSeleccionar.AppendLine(" 			SUM(Base_B) 								AS Base_B,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Costos) 								AS Costos,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Costo_B)								AS Costo_B,")
                    loComandoSeleccionar.AppendLine(" 			SUM(Ganancias)								AS Ganancias,")
                    loComandoSeleccionar.AppendLine(" 			AVG(Ganancia_B)								AS Ganancia_B")
                    loComandoSeleccionar.AppendLine("FROM #Temp ")
                    loComandoSeleccionar.AppendLine("GROUP BY DATEPART(YEAR,Fec_Ini), DATEPART(MONTH,Fec_Ini)")
            End Select

			loComandoSeleccionar.AppendLine("ORDER BY 	DATEPART(YEAR,Fec_Ini) , DATEPART(MONTH,Fec_Ini) ")
			
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
            
            
    '******************************************************************************************
	' Inicio Se Procesa manualmetne los datos
	'******************************************************************************************

			'Tabla con las listas desplegables
			Dim loTabla As New DataTable("curReportes")
			Dim loColumna As DataColumn 
			
			loColumna = New DataColumn("Año", getType(Integer))
			loTabla.Columns.Add(loColumna)
			
			loColumna = New DataColumn("Mes", getType(Integer))
			loTabla.Columns.Add(loColumna)
			
			loColumna = New DataColumn("Ventas", getType(Decimal))
			loTabla.Columns.Add(loColumna)
			
			loColumna = New DataColumn("Base_B", getType(Decimal))
			loTabla.Columns.Add(loColumna)
			
			loColumna = New DataColumn("Costos", getType(Decimal))
			loTabla.Columns.Add(loColumna)
			
			loColumna = New DataColumn("Costo_B", getType(Decimal))
			loTabla.Columns.Add(loColumna)
			
			loColumna = New DataColumn("Ganancias", getType(Decimal))
			loTabla.Columns.Add(loColumna)
			
			loColumna = New DataColumn("Ganancia_B", getType(Decimal))
			loTabla.Columns.Add(loColumna)
			
			Dim llVacio As Boolean = False
			Dim loNuevaFila As DataRow
			Dim lnTotalFilas AS Integer = laDatosReporte.Tables(0).Rows.Count

			If lnTotalFilas > 0 Then

				Dim Contador As Integer = 0	
				Dim Contador2 As Integer = 0	


				For i As  Integer = laDatosReporte.Tables(0).Rows(0).Item("Año_Inicio") To laDatosReporte.Tables(0).Rows(0).Item("Año_Fin")
					For j As  Integer = 1 To 12
						

						IF (i = laDatosReporte.Tables(0).Rows(0).Item("Año_Inicio")) AND (j <= laDatosReporte.Tables(0).Rows(0).Item("Mes_Inicio")-1) Then 
							
							Contador2 = 0
							
						Else
							
							IF (i = laDatosReporte.Tables(0).Rows(0).Item("Año_Fin")) AND (j > laDatosReporte.Tables(0).Rows(0).Item("Mes_Fin")) Then 

								Contador2 = 0

							Else
								loNuevaFila = loTabla.NewRow()
								loTabla.Rows.Add(loNuevaFila)
								
									
									loNuevaFila.Item("Año")				= i
									loNuevaFila.Item("Mes")				= (j)
									loNuevaFila.Item("Ventas")			= 0.0
									loNuevaFila.Item("Base_B")			= 0.0
									loNuevaFila.Item("Costos")			= 0.0
									loNuevaFila.Item("Costo_B")			= 0.0
									loNuevaFila.Item("Ganancias")		= 0.0
									loNuevaFila.Item("Ganancia_B")		= 0.0

								
								loTabla.AcceptChanges()
							
							End If
							
						End if
						
					Next
				Next
						
				Dim lnAux As Integer = 0
				For Each loRenglonActual As DataRow In laDatosReporte.Tables(0).Rows

					Dim loRenglon As DataRow
					Dim lcAxuAño As String = loRenglonActual.Item("Año").ToString
					Dim lcAxuMes As String = loRenglonActual.Item("Mes").ToString
							
					
					'If LnAux = 0  Then
						
					'	LnAux = LnAux + 1
					'	Continue For
						
					'end if 	
					
					'If LnAux > 0 Then

						loRenglon = loTabla.Select("Año = " & lcAxuAño & " AND Mes = " & lcAxuMes) (0)
						
						loRenglon("Año")			= loRenglonActual("Año")
						loRenglon("Mes")			= loRenglonActual("Mes")
						loRenglon("Ventas")			= loRenglonActual("Ventas")
						loRenglon("Base_B")			= loRenglonActual("Base_B")
						loRenglon("Costos")			= loRenglonActual("Costos")
						loRenglon("Costo_B")		= loRenglonActual("Costo_B")
						loRenglon("Ganancias")		= loRenglonActual("Ganancias")
						loRenglon("Ganancia_B")		= loRenglonActual("Ganancia_B")
					
					'End If
					
					'LnAux = LnAux + 1

				Next
			
			Else
				
					llVacio = True
				
			End If

	'******************************************************************************************
	' Fin Se Procesa manualmetne los datos
	'******************************************************************************************
                        
		 If	 llVacio = False

		 
			Dim loTablaRellenada As New DataSet()
			loTablaRellenada.Tables.Add(loTabla)


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grMargen_gEmpresa", loTablaRellenada)

			CType(loObjetoReporte.ReportDefinition.ReportObjects("text2"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(0).Rows(0).Item("Año_Inicio").ToString
			CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(0).Rows(0).Item("Año_Fin").ToString
			'CType(loObjetoReporte.ReportDefinition.ReportObjects("text2"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "[Sin Fecha]"
			'CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = "[Sin Fecha]"


            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrMargen_gEmpresa.ReportSource = loObjetoReporte
            
        Else
			
			'-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------
					
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
				
        End If

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
' CMS: 06/05/10: Programacion inicial														'
'-------------------------------------------------------------------------------------------'
' RJG: 29/06/12: Se corrigió bug cuando el reporte solo devuelve una fila.					'
'-------------------------------------------------------------------------------------------'
