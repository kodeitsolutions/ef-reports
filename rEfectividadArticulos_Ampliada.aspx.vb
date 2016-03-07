'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEfectividadArticulos_Ampliada"
'-------------------------------------------------------------------------------------------'
Partial Class rEfectividadArticulos_Ampliada
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Private Const KN_LONGITUD_CODIGO_0 = 2
	Private Const KN_LONGITUD_CODIGO_1 = 4
	Private Const KN_LONGITUD_CODIGO_2 = 6
	Private Const KN_LONGITUD_CODIGO_3 = 7
	
    Private Dim lnEfectividadBaja As Decimal 
    Private Dim lnEfectividadAlta As Decimal 
    Private Dim lnGananciaBaja As Decimal
    Private Dim lnCompraAdicionalSugerida As Decimal
    Private Dim lnDiasEvaluarInventario As Decimal
    
    Private Dim ldFechaInicio As Date
    Private Dim ldFechaFin As Date

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

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
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            
            Dim lcParametro11Desde As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(11)).Trim().ToUpper()
            Dim lcParametro12Desde As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(12)).Trim()
            
            Dim lnParametro13Desde As Decimal = CDec(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lnParametro14Desde As Decimal = CDec(cusAplicacion.goReportes.paParametrosIniciales(14))
            Dim lnParametro15Desde As Decimal = CDec(cusAplicacion.goReportes.paParametrosIniciales(15))
            Dim lnParametro16Desde As Decimal = CDec(cusAplicacion.goReportes.paParametrosIniciales(16))
            Dim lnParametro17Desde As Decimal = CDec(cusAplicacion.goReportes.paParametrosIniciales(17))

			Dim lcPrecio AS String = "Precio1"
			Select Case lcParametro12Desde
				Case "1"
					lcPrecio = "Precio1"
				Case "2"
					lcPrecio = "Precio2"
				Case "3"
					lcPrecio = "Precio3"
				Case "4"
					lcPrecio = "Precio4"
				Case "5"
					lcPrecio = "Precio5"
			End Select
			
			'Parámetros para Excel
			lnEfectividadBaja 			= lnParametro13Desde
			lnEfectividadAlta			= lnParametro14Desde
			lnGananciaBaja				= lnParametro15Desde
			lnCompraAdicionalSugerida	= lnParametro16Desde
			lnDiasEvaluarInventario		= IIf(lnParametro17Desde>0, lnParametro17Desde, 30)
			
			ldFechaInicio	= CDate(cusAplicacion.goReportes.paParametrosIniciales(0))
			ldFechaFin		= CDate(cusAplicacion.goReportes.paParametrosFinales(0))
			
			
			
			'Inicio del reporte (SQL)			
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

			

            loComandoSeleccionar.AppendLine("SELECT		    Articulos.Cod_Art						AS Cod_Art,		")
            loComandoSeleccionar.AppendLine("           	Articulos.Nom_Art						AS Nom_Art,		")
            loComandoSeleccionar.AppendLine("           	Articulos." & lcPrecio & "				AS Precio,		")
            loComandoSeleccionar.AppendLine("           	Articulos.Cos_Pro1						AS Cos_Pro,		")
            loComandoSeleccionar.AppendLine("           	SUM(COALESCE(Renglones.Can_Art1, 0))	AS Can_Art1,	")
            loComandoSeleccionar.AppendLine("           	SUM(COALESCE(Renglones.Mon_Net, 0))		AS Mon_Net,		")
            loComandoSeleccionar.AppendLine("           	Articulos.Exi_Act1						AS Exi_Act1,	")
            loComandoSeleccionar.AppendLine("           	Articulos.Exi_Por1						AS Exi_Por1,	")
            loComandoSeleccionar.AppendLine("           	Articulos.Cod_Mar						AS Cod_Mar		")
            loComandoSeleccionar.AppendLine("INTO		#tmpCotizaciones")
            loComandoSeleccionar.AppendLine("FROM		Articulos ")
        'If	(lcParametro11Desde = "SI") Then
            'loComandoSeleccionar.AppendLine("	JOIN (")
        'Else
            loComandoSeleccionar.AppendLine("	LEFT JOIN (")
        'End If
            loComandoSeleccionar.AppendLine("				SELECT		Renglones_Cotizaciones.Cod_Art, ")
            loComandoSeleccionar.AppendLine("							Renglones_Cotizaciones.Can_Art1, ")
            loComandoSeleccionar.AppendLine("							Renglones_Cotizaciones.Mon_Net")
            loComandoSeleccionar.AppendLine(" 				FROM		Cotizaciones")
            loComandoSeleccionar.AppendLine("					JOIN	Renglones_Cotizaciones ")
            loComandoSeleccionar.AppendLine("						ON	Renglones_Cotizaciones.Documento = Cotizaciones.Documento	")
            loComandoSeleccionar.AppendLine("				WHERE		Cotizaciones.Status IN ('Confirmado', 'Afectado', 'Procesado')")
			loComandoSeleccionar.AppendLine(" 						AND Cotizaciones.Fec_Ini BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 							AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 						AND Cotizaciones.Cod_Cli BETWEEN " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine(" 							AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine(" 						AND Cotizaciones.Cod_Ven BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 							AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 						AND Cotizaciones.Cod_Suc BETWEEN " & lcParametro9Desde)
			loComandoSeleccionar.AppendLine(" 							AND " & lcParametro9Hasta)
			loComandoSeleccionar.AppendLine(" 						AND Cotizaciones.Cod_Rev BETWEEN " & lcParametro10Desde)
			loComandoSeleccionar.AppendLine(" 							AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("			) AS Renglones")
            loComandoSeleccionar.AppendLine("		ON	Renglones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE					Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 							AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			        	AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			        	    AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			        	AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			        	    AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			        	AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			        	    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			        	AND Articulos.cod_Mar BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			        	    AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Articulos.Cod_Art,")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art,")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Act1,")
            loComandoSeleccionar.AppendLine("			Articulos.Exi_Por1,")
            loComandoSeleccionar.AppendLine("			Articulos." & lcPrecio & ",")
            loComandoSeleccionar.AppendLine("			Articulos.Cos_Pro1,")
            loComandoSeleccionar.AppendLine("			Articulos.Cod_Mar")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT     ")
            loComandoSeleccionar.AppendLine("        	Renglones_Facturas.Cod_Art,")
            loComandoSeleccionar.AppendLine("        	SUM(Renglones_Facturas.Can_Art1)	AS Can_Art1,")
            loComandoSeleccionar.AppendLine("        	SUM(Renglones_Facturas.Mon_Net) 	AS Mon_Net,")
            loComandoSeleccionar.AppendLine("        	MAX(Facturas.Fec_Ini)				AS Fecha_Ultima_Venta")
            loComandoSeleccionar.AppendLine("INTO	#tmpFacturas")
            loComandoSeleccionar.AppendLine("FROM	Facturas")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            loComandoSeleccionar.AppendLine(" 			                AND Renglones_Facturas.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("JOIN	Articulos ON Renglones_Facturas.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Mar BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("WHERE  	Facturas.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.cod_Suc BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Rev BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine("SELECT		Renglones_Compras.Cod_Art			AS Cod_Art,")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Compras.Can_Art1)		AS Can_Art1,")
            loComandoSeleccionar.AppendLine("			Renglones_Compras.Precio1			AS Precio1,")  
            loComandoSeleccionar.AppendLine("			MAX(Compras.Fec_Ini)				AS Fecha_Ultima_Compra,")
            loComandoSeleccionar.AppendLine("        	Compras.Cod_Mon						AS Cod_Mon_Compra")
            loComandoSeleccionar.AppendLine("INTO	#tmpCompras")
            loComandoSeleccionar.AppendLine("FROM	Compras")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento")
            loComandoSeleccionar.AppendLine("                           AND Renglones_Compras.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("JOIN	Articulos ON Renglones_Compras.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.cod_Mar BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("WHERE  Compras.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine("           AND Compras.Fec_Ini < " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Compras.Tip_Ori	<>	'Recepciones' ")
            loComandoSeleccionar.AppendLine("GROUP BY Renglones_Compras.Cod_Art,Renglones_Compras.Precio1,Compras.Cod_Mon")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("UNION ALL")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Renglones_Recepciones.Cod_Art		AS Cod_Art,")
            loComandoSeleccionar.AppendLine("           SUM(Renglones_Recepciones.Can_Art1) As Can_Art1,")
            loComandoSeleccionar.AppendLine("           Renglones_Recepciones.Precio1		AS Precio1,")
            loComandoSeleccionar.AppendLine("           MAX(Recepciones.Fec_Ini)			AS Fecha_Ultima_Compra,")
            loComandoSeleccionar.AppendLine("        	Recepciones.Cod_Mon					AS Cod_Mon_Compra")
            loComandoSeleccionar.AppendLine("FROM	Recepciones")
            ' loComandoSeleccionar.AppendLine("JOIN	Proveedores ON Proveedores.Cod_Pro = Recepciones.Cod_Pro")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Recepciones ON Renglones_Recepciones.Documento = Recepciones.Documento")
            loComandoSeleccionar.AppendLine("                           AND Renglones_Recepciones.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("JOIN	Articulos ON Renglones_Recepciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Mar BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("WHERE  Recepciones.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine("           AND Recepciones.Fec_Ini < " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY Renglones_Recepciones.Cod_Art,Renglones_Recepciones.Precio1,Recepciones.Cod_Mon")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("UNION ALL")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Renglones_Ajustes.Cod_Art			AS Cod_Art,")
            loComandoSeleccionar.AppendLine("           SUM(Renglones_Ajustes.Can_Art1)		AS Can_Art1,")
            loComandoSeleccionar.AppendLine("           Renglones_Ajustes.Cos_Ult1			AS Precio1,") 
            loComandoSeleccionar.AppendLine("           MAX(Ajustes.Fec_Ini)				AS Fecha_Ultima_Compra,")
            loComandoSeleccionar.AppendLine("        	Ajustes.Cod_Mon						AS Cod_Mon_Compra")
            loComandoSeleccionar.AppendLine("FROM	Ajustes")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento")
            loComandoSeleccionar.AppendLine("                           AND Renglones_Ajustes.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 		                    AND	Renglones_Ajustes.Tipo	IN	('Entrada') ")
            loComandoSeleccionar.AppendLine("JOIN	Articulos ON Renglones_Ajustes.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			                AND Articulos.cod_Mar BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			                AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("WHERE  Ajustes.Status IN ('Confirmado')")
            loComandoSeleccionar.AppendLine("           AND Ajustes.Fec_Ini < " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY Renglones_Ajustes.Cod_Art,Renglones_Ajustes.Cos_Ult1,Ajustes.Cod_Mon")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine("SELECT		  #tmpTemporal.Cod_Art,")
            loComandoSeleccionar.AppendLine("             #tmpTemporal.Can_Art1,")
            loComandoSeleccionar.AppendLine("             #tmpTemporal.Precio1,") 
            loComandoSeleccionar.AppendLine("             #tmpTemporal.Fecha_Ultima_Compra,")
            loComandoSeleccionar.AppendLine("             #tmpTemporal.Cod_Mon_Compra")
            loComandoSeleccionar.AppendLine("INTO #tmpUltimasCompras")
            loComandoSeleccionar.AppendLine("FROM #tmpCompras AS #tmpTemporal")
            loComandoSeleccionar.AppendLine("	INNER JOIN (SELECT Cod_Art,MAX(Fecha_Ultima_Compra) As Fecha_Maxima FROM #tmpCompras GROUP BY Cod_Art) AS Tabla")
            loComandoSeleccionar.AppendLine("		ON #tmpTemporal.Fecha_Ultima_Compra = Tabla.Fecha_Maxima AND #tmpTemporal.Cod_Art = Tabla.Cod_Art")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Art Asc")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT     ")
            loComandoSeleccionar.AppendLine("			#tmpCotizaciones.Cod_Art				AS Cod_Art,")
            loComandoSeleccionar.AppendLine("			#tmpCotizaciones.Nom_Art				AS Nom_Art,")    
            loComandoSeleccionar.AppendLine("			#tmpCotizaciones.Precio					AS Precio,")    
            loComandoSeleccionar.AppendLine("			#tmpCotizaciones.Cos_Pro				AS Cos_Pro,")    
            loComandoSeleccionar.AppendLine("			#tmpCotizaciones.Can_Art1				AS Can_Art1_Cot,")
            loComandoSeleccionar.AppendLine("			COALESCE(#tmpFacturas.Can_Art1,0)		AS Can_Art1_Fac,")
            loComandoSeleccionar.AppendLine("			CASE WHEN #tmpCotizaciones.Can_Art1>0")
            loComandoSeleccionar.AppendLine("				THEN	COALESCE(#tmpFacturas.Can_Art1,0)*100/#tmpCotizaciones.Can_Art1")
            loComandoSeleccionar.AppendLine("				ELSE	-1")
            loComandoSeleccionar.AppendLine("			END										AS Efic1,")
            loComandoSeleccionar.AppendLine("			#tmpCotizaciones.Exi_Act1				AS Exi_Act1,")
            loComandoSeleccionar.AppendLine("			#tmpCotizaciones.Exi_Por1				AS Exi_Por1,")
            loComandoSeleccionar.AppendLine("			#tmpCotizaciones.Cod_Mar				AS Cod_Mar,")
            loComandoSeleccionar.AppendLine("			#tmpFacturas.Fecha_Ultima_Venta			AS Fecha_Ultima_Venta")
            loComandoSeleccionar.AppendLine("INTO		#tmpFinal")
            loComandoSeleccionar.AppendLine("FROM		#tmpCotizaciones")
            loComandoSeleccionar.AppendLine("LEFT JOIN	#tmpFacturas ON #tmpFacturas.Cod_Art = #tmpCotizaciones.Cod_Art")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT		#tmpFinal.Cod_Art,")
            loComandoSeleccionar.AppendLine("			#tmpFinal.Nom_Art,")    
            loComandoSeleccionar.AppendLine("			#tmpFinal.Precio,")    
            loComandoSeleccionar.AppendLine("			#tmpFinal.Cos_Pro,")    
            loComandoSeleccionar.AppendLine("			CASE WHEN (COALESCE(#tmpUltimasCompras.Precio1,0)>0) ")    
            loComandoSeleccionar.AppendLine("				THEN (#tmpFinal.Precio - COALESCE(#tmpUltimasCompras.Precio1,0))/COALESCE(#tmpUltimasCompras.Precio1,0)*100")    
            loComandoSeleccionar.AppendLine("				ELSE 100")    
            loComandoSeleccionar.AppendLine("			END		AS Ganancia,")    
            loComandoSeleccionar.AppendLine("			#tmpFinal.Can_Art1_Cot,")
            loComandoSeleccionar.AppendLine("			#tmpFinal.Can_Art1_Fac,")
            loComandoSeleccionar.AppendLine("			#tmpFinal.Efic1,")
            loComandoSeleccionar.AppendLine("			#tmpFinal.Exi_Act1,")
            loComandoSeleccionar.AppendLine("			#tmpFinal.Exi_Por1,")
            loComandoSeleccionar.AppendLine("			#tmpFinal.Cod_Mar,")
            loComandoSeleccionar.AppendLine("			#tmpFinal.Fecha_Ultima_Venta,")
            loComandoSeleccionar.AppendLine("           COALESCE(#tmpUltimasCompras.Can_Art1,0)		        AS Cantidad_Comprada,")
            loComandoSeleccionar.AppendLine("           COALESCE(#tmpUltimasCompras.Precio1,0)		        AS Costo_Ultimo,")
            loComandoSeleccionar.AppendLine("           COALESCE(#tmpUltimasCompras.Fecha_Ultima_Compra,0)	AS Fecha_Ultima_Compra, ")
            loComandoSeleccionar.AppendLine("           COALESCE(#tmpUltimasCompras.Cod_Mon_Compra,'')	AS Cod_Mon_Compra ")
            loComandoSeleccionar.AppendLine("FROM		#tmpFinal            ")
            loComandoSeleccionar.AppendLine("LEFT JOIN  #tmpUltimasCompras  ON #tmpUltimasCompras.Cod_Art = #tmpFinal.Cod_Art   ")
        If	(lcParametro11Desde = "SI") Then
            loComandoSeleccionar.AppendLine("WHERE		ABS(#tmpFinal.Can_Art1_Cot) ")
            loComandoSeleccionar.AppendLine("		+	ABS(#tmpFinal.Can_Art1_Fac) ")
            loComandoSeleccionar.AppendLine("		+	ABS(#tmpFinal.Exi_Act1) ")
            loComandoSeleccionar.AppendLine("		+	ABS(#tmpFinal.Exi_Por1) ")
            loComandoSeleccionar.AppendLine("		+	ABS(COALESCE(#tmpUltimasCompras.Can_Art1,0)) > 0")
         End If
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCompras")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpFinal")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpUltimasCompras")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCotizaciones ")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpFacturas ")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes", 9000)

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEfectividadArticulos_Ampliada", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrEfectividadArticulos_Ampliada.ReportSource = loObjetoReporte
            

            'Selección de opcion por excel (Microsoft Excel - xls)
            If (Me.Request.QueryString("salida").ToLower = "xls") Then
                
                Me.Server.ScriptTimeout = 9000 

                ' Ruta donde se creara temporalmente el archivo
                Dim lcFileName As String = Server.MapPath("~\Administrativo\Temporales\rEfectividadArticulos_Ampliada_" & Guid.NewGuid().ToString("N") & ".xls")
                ' Se exporta para crear el archivo temporal

                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, lcFileName)

                ' Se modifica el contenido del archivo
                Me.mGenerarArchivoExcel(lcFileName, laDatosReporte.Tables(0), "")

                ' Se coloca en la respuesta para decargar
                Me.Response.Clear()
                Me.Response.Buffer = True 
                Me.Response.AppendHeader("content-disposition", "attachment; filename=rEfectividadArticulos_Ampliada.xls")
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

	Private Sub mGenerarArchivoExcel(ByVal lcNombreArchivo As String, ByVal loDatos As DataTable, ByVal lcParametrosReporte As String)
		
		Dim lnDecimales As Integer = goOpciones.pnDecimalesParaMonto
		Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
		
		Dim lcFormatoMontos As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimales)
		
		Dim lcFormatoCantidad As String 
		If (lnDecimalesCantidad > 0) Then 
			lcFormatoCantidad = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesCantidad)
		Else
			lcFormatoCantidad = "###,###,###,###,##0"
		End If
		
		Dim loFormatoFecha As IFormatProvider 
		loFormatoFecha = System.Globalization.CultureInfo.CreateSpecificCulture("es-VE").GetFormat(GetType(Date))
		
		Dim lcFechaInicio	As String = CDate(ldFechaInicio).ToString("dd-MMM-yyyy", loFormatoFecha)
		Dim lcFechaFin		As String = CDate(ldFechaFin).ToString("dd-MMM-yyyy", loFormatoFecha)

	 '******************************************************************'
	 ' Declaración de objetos de excel: IMPORTANTE liberar recursos al	'
	 ' final usando el GARBAGE COLLECTOR y ReleaseComObject.			'
	 '******************************************************************'
		Dim loExcel		As Excel.Application	= Nothing
		Dim laLibros	As Excel.Workbooks		= Nothing
		Dim loLibro		As Excel.Workbook		= Nothing
        Dim loHoja		As Excel.Worksheet		= Nothing
		Dim loCeldas	As Excel.Range			= Nothing
		Dim loRango		As Excel.Range			= Nothing
		
		Dim loFilas		As Excel.Range			= Nothing
		Dim loColumnas	As Excel.Range			= Nothing
		Dim loFormas	As Excel.Shapes			= Nothing
		Dim loImagen	As Excel.Shape			= Nothing
		Dim loFuente	As Excel.Font			= Nothing
		
		
        Try
        
        ' Se inicializa el objeto de aplicacion excel
            loExcel = New Excel.Application()
            loExcel.Visible = False
            loExcel.DisplayAlerts = False 

        ' Crea un nuevo libro de excel y activa la primera hoja
            laLibros = loExcel.Workbooks
            'loLibro = laLibros.Add()
            
            'Dim lcPlantilla As String = HttpContext.Current.Server.MapPath("~/Administrativo/Complementos/plantilla.xls")
            'System.IO.File.Copy(lcPlantilla, lcNombreArchivo)
            loLibro = laLibros.Open(lcNombreArchivo)
            
            loHoja = loLibro.Worksheets(1)
            loHoja.Activate()

            loExcel.Calculation = Excel.XlCalculation.xlCalculationManual 
            loExcel.CalculateBeforeSave = True

		' Formato por defecto de todas las celdas			
			loCeldas = loHoja.Range("A1:IV65536")
            'loCeldas = loHoja.Cells
			loCeldas.Clear()
            loFuente = loCeldas.Font
            loFuente.Size = 8
            loFuente.Name = "Tahoma"


		 '******************************************************************'
		 ' Encabezado de la hoja											'
		 '******************************************************************'
			'Dim lcLogo As String = goEmpresa.pcUrlLogo 
			'lcLogo = HttpContext.Current.Server.MapPath(lcLogo)
			'loFormas = loHoja.Shapes

			'loFormas.AddPicture(lcLogo, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 1, 1, 60, 60)
			
            loRango = loHoja.Range("A1")
            loRango.Value = cusAplicacion.goEmpresa.pcNombre
            loFuente = loRango.Font
            loFuente.Size = 16
            loFuente.Bold = True
            
            loRango = loHoja.Range("A2")
            loRango.Value = cusAplicacion.goEmpresa.pcRifEmpresa
            loFuente = loRango.Font
            loFuente.Size = 10
            loFuente.Bold = True

            loRango = loHoja.Range("B8:F8")
            loRango.Select()
            loRango.MergeCells = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.Value = "Efectividad por Producto"
            loFuente = loRango.Font
            loFuente.Size = 12
            loFuente.Bold = True
            loFuente.Color = &HFF0000 'Azul

            loRango = loHoja.Range("B9:F9")
            loRango.Select()
            loRango.MergeCells = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.Value = "MOVIMIENTO DEL PERIODO COMPRENDIDO DESDE " & lcFechaInicio & "  HASTA " & lcFechaFin
            loFuente = loRango.Font
            loFuente.Size = 10
            loFuente.Bold = True

            ' Fecha y hora de creacion
			Dim ldFecha As DateTime = Date.Now()
			loRango = loHoja.Range("O1")
			loRango.NumberFormat = "@"
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
			loRango.Value = ldFecha.ToString("dd/MM/yyyy")
			loRango = loHoja.Range("O2")
			loRango.NumberFormat = "@"
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
			loRango.Value = ldFecha.ToString("hh:mm:ss tt")

            ' Parametros del reporte
            'loRango = loHoja.Range("A9:O9")
            'loRango.Select()
            'loRango.MergeCells = True
            'loRango.Value = lcParametrosReporte
            'loRango.WrapText = True
            'loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify


			Dim lnFilaActual As Integer = 10

		 '******************************************************************'
		 ' Datos del Encabezado												'
		 '******************************************************************'
			
			loRango = loHoja.Range("B" & lnFilaActual)
			loRango.Value = "Artículo"
			
			loRango = loHoja.Range("C" & lnFilaActual)
			loRango.Value = "Descripción"
			
			loRango = loHoja.Range("D" & lnFilaActual)
			loRango.Value = "Cantidad"  & Strings.Chr(10) & "Cotizado"
			
			loRango = loHoja.Range("E" & lnFilaActual)
			loRango.Value = "Cantidad"  & Strings.Chr(10) & "Facturado"
			
			loRango = loHoja.Range("F" & lnFilaActual)
			loRango.Value = "Efectividad"
			
			loRango = loHoja.Range("G" & lnFilaActual)
			loRango.Value = "Existencia"  & Strings.Chr(10) & "Actual"
			
			loRango = loHoja.Range("H" & lnFilaActual)
			loRango.Value = "Existencia"  & Strings.Chr(10) & "por Llegar"
			
			loRango = loHoja.Range("I" & lnFilaActual)
			loRango.Value = "Fecha Últ."  & Strings.Chr(10) & "Compra"
			
			loRango = loHoja.Range("J" & lnFilaActual)
			loRango.Value = "Cantidad"  & Strings.Chr(10) & "Últ. Compra"
			
			loRango = loHoja.Range("K" & lnFilaActual)
			loRango.Value = "Costo"
			
			loRango = loHoja.Range("L" & lnFilaActual)
			loRango.Value = "Precio"
			
			loRango = loHoja.Range("M" & lnFilaActual)
			loRango.Value = "%"
			
			loRango = loHoja.Range("N" & lnFilaActual)
			loRango.Value = "Comprar"  & Strings.Chr(10) & "Unidades"
			
			loRango = loHoja.Range("O" & lnFilaActual)
			loRango.Value = "Total"  & Strings.Chr(10) & "Comprar"
			
			'Campos para evaluación de productos débiles (HUESO)
			
			'Marca: Q
			loRango = loHoja.Range("Q" & lnFilaActual)
			loRango.Value = "MARCA"
			
			'Codigo 0: R
			loRango = loHoja.Range("R" & lnFilaActual)
			loRango.Value = "COD 0"
			
			'Codigo 1: S
			loRango = loHoja.Range("S" & lnFilaActual)
			loRango.Value = "COD 1"
			
			'Codigo 2: T
			loRango = loHoja.Range("T" & lnFilaActual)
			loRango.Value = "COD 2"
			
			'Codigo 3: U
			loRango = loHoja.Range("U" & lnFilaActual)
			loRango.Value = "COD 3"
			
			'Datos: V
			loRango = loHoja.Range("V" & lnFilaActual)
			loRango.Value = "DATOS"
			
			'Inventario Actual $: W
			loRango = loHoja.Range("W" & lnFilaActual)
			loRango.Value = "INV $"
			
			'Inventario por Llegar $: X
			loRango = loHoja.Range("X" & lnFilaActual)
			loRango.Value = "INV POR" & Strings.Chr(10) & "LLEGAR"
	
			'Débil (HUESO): Y
			loRango = loHoja.Range("Y" & lnFilaActual)
			loRango.Value = "DÉBIL"
				
			'Inventario Vendido $: Z
			loRango = loHoja.Range("Z" & lnFilaActual)
			loRango.Value = "INV." & Strings.Chr(10) & "VENDIDO "
			
			'Débil (HUESO) por Marca: AA
			loRango = loHoja.Range("AA" & lnFilaActual)
			loRango.Value = "DEBIL" & Strings.Chr(10) & "POR " & Strings.Chr(10) & "MARCA"

			
			'Encabezado 1
			loRango = loHoja.Range("B" & lnFilaActual & ":O" & lnFilaActual)
			loFuente = loRango.Font
			loFuente.Bold = True
			loFuente.Color = &HFFFFFF  'Blanco
			loRango.Interior.Color = &H000000  'Negro
			loRango.Interior.Pattern = Excel.XlPattern.xlPatternSolid 
			
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
			loRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
			
			'Encabezado 2
			loRango = loHoja.Range("Q" & lnFilaActual & ":AA" & lnFilaActual)
			loFuente = loRango.Font
			loFuente.Bold = True
			loFuente.Color = &H000000  'Blanco
			loRango.Interior.Color = &H00FFFF  'Amarillo
			loRango.Interior.Pattern = Excel.XlPattern.xlPatternSolid 
			
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
			loRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)

			With loRango.Borders(Excel.XlBordersIndex.xlInsideVertical)
				.LineStyle = Excel.XlLineStyle.xlContinuous
				.Weight = Excel.XlBorderWeight.xlThin
			End With
		

        '***************************************************************************
        ' Inicio Formato tabla (colocar antes/fuera del FOR para optimizar)
	    '***************************************************************************
            Dim lnInicioTabla As Integer = lnFilaActual + 1
            Dim lnFinTabla As Integer = lnFilaActual + loDatos.Rows.Count
            
			'Artículo: B
			'Descripción: C
		    loRango = loHoja.Range("B" & lnInicioTabla & ":C" & lnFinTabla)
			loRango.NumberFormat = "@"
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

			'Cotizado: D
			'Facturado: E
		    loRango = loHoja.Range("D" & lnInicioTabla & ":D" & lnFinTabla)
			loRango.NumberFormat = lcFormatoCantidad
			
			'Stock Actual: G
		    loRango = loHoja.Range("G" & lnInicioTabla & ":G" & lnFinTabla)
			loRango.NumberFormat = lcFormatoCantidad
			loFuente = loRango.Font
			loFuente.Bold = True

            'Stock X Llegar: H
		    loRango = loHoja.Range("H" & lnInicioTabla & ":H" & lnFinTabla)
			loRango.NumberFormat = lcFormatoCantidad

		    'Fecha Última Compra: I
		    loRango = loHoja.Range("I" & lnInicioTabla & ":I" & lnFinTabla)
			loRango.NumberFormat = "[$-200A]dd-mmm-yyyy;@"		'--> [$-200A]= Español; [$-409A]= Inglés, etc...
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

			'Unidades Últ. Compra: J
			'Costo: K
		    loRango = loHoja.Range("J" & lnInicioTabla & ":K" & lnFinTabla)
			loRango.NumberFormat = lcFormatoMontos

			'Precio: L
		    loRango = loHoja.Range("L" & lnInicioTabla & ":L" & lnFinTabla)
			loRango.NumberFormat = lcFormatoMontos
			loFuente = loRango.Font
			loFuente.Bold = True

			'Ganancia: M
		    loRango = loHoja.Range("M" & lnInicioTabla & ":M" & lnFinTabla)
			loRango.NumberFormat = lcFormatoMontos
			loFuente = loRango.Font
			loFuente.Color = &H808080 'Gris 50%

			'Comprar Unidades: N
		    loRango = loHoja.Range("N" & lnInicioTabla & ":N" & lnFinTabla)
			loRango.NumberFormat = lcFormatoCantidad

			'Total Comprar: O
		    loRango = loHoja.Range("O" & lnInicioTabla & ":O" & lnFinTabla)
			loRango.NumberFormat = lcFormatoMontos

			'Lineas horizontales
		    loRango = loHoja.Range("B" & lnInicioTabla & ":O" & lnFinTabla)
			With loRango.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
				.LineStyle = Excel.XlLineStyle.xlDash
				.Weight = Excel.XlBorderWeight.xlHairline
				.Color = &H808080 'Gris 50%
			End With

			'Marca: Q
		    loRango = loHoja.Range("Q" & lnInicioTabla & ":Q" & lnFinTabla)
			loRango.NumberFormat = "@"

            'Datos: V
			'Inventario Actual $: W
			'Inventario por Llegar $: X
			'Débil (HUESO): Y
			'Inventario Vendido $: Z
			'Hueso por Marca: AA
		    loRango = loHoja.Range("V" & lnInicioTabla & ":AA" & lnFinTabla)
			loRango.NumberFormat = lcFormatoMontos

			'Linea completa 
			loRango = loHoja.Range("Q" & lnInicioTabla & ":AA" & lnFinTabla) 
			loRango.Interior.ColorIndex = 40 'Gris claro
			
		'***************************************************************************
        ' Fin Formato tabla
		'***************************************************************************

			
			Dim lnFilaInicio As Integer  = lnFilaActual
			For lnRenglon As Integer = 0 To loDatos.Rows.Count - 1
				Dim loRenglon As DataRow = loDatos.Rows(lnRenglon)
				
				lnFilaActual += 1

				'Artículo: B
				loRango = loCeldas(lnFilaActual, 2) 
				loRango.Value = CStr(loRenglon("Cod_Art")).Trim()
				
				'Descripción: C
				loRango = loCeldas(lnFilaActual, 3) 
				loRango.Value = CStr(loRenglon("Nom_Art")).Trim()
				
				'Cotizado: D
				loRango = loCeldas(lnFilaActual, 4) 
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Can_Art1_Cot")), lnDecimalesCantidad)
				
				'Facturado: E
				loRango = loCeldas(lnFilaActual, 5) 
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Can_Art1_Fac")), lnDecimalesCantidad)
				
				'Eficiencia %: F
				Dim lnEficiencia As Decimal = CDec(loRenglon("Efic1"))
				loRango = loCeldas(lnFilaActual, 6) 
                If (lnEficiencia > 0) Then
					loRango.NumberFormat = "0.00"
					loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Efic1")), 2)
				Else
					loRango.NumberFormat = "@"
					loRango.Value = "-"
					loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
				End If

				'Stock Actual: G
				loRango = loCeldas(lnFilaActual, 7) 
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Exi_Act1")), lnDecimalesCantidad)
				
				'Stock X Llegar: H
				loRango = loCeldas(lnFilaActual, 8) 
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Exi_Por1")), lnDecimalesCantidad)
				
				'Fecha Última Compra: I
				loRango = loCeldas(lnFilaActual, 9) 
				loRango.Value = CDate(loRenglon("Fecha_Ultima_Compra")) '.ToString("dd/MM/yyyy")

				'Unidades Últ. Compra: J
				loRango = loCeldas(lnFilaActual, 10) 
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Cantidad_Comprada")), lnDecimalesCantidad)
				
				'Costo: K
				loRango = loCeldas(lnFilaActual, 11) 
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Costo_Ultimo")), lnDecimales)
				
				'Precio: L
				loRango = loCeldas(lnFilaActual, 12) 
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Precio")), lnDecimales)
				
				'Ganancia: M
				loRango = loCeldas(lnFilaActual, 13) 
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Ganancia")), lnDecimales)
				
				'Comprar Unidades: N
				loRango = loCeldas(lnFilaActual, 14) 
				loRango.Value = "=ROUND(IF(E" & lnFilaActual & ">0, IF(G" & lnFilaActual & "<E" & lnFilaActual & ", (E" & lnFilaActual & "-G" & lnFilaActual & ")*(100+$AA$5)/100, 0), 0), " & lnDecimalesCantidad & ")" 
				
				'Total Comprar: O
				loRango = loCeldas(lnFilaActual, 15) 
				loRango.Value = "=ROUND(IF(N" & lnFilaActual & ">0, N" & lnFilaActual & "*K" & lnFilaActual & ",0), " & lnDecimales & ")"
				
				'***************************************************************
				' Cálculo de productos Débiles (Hueso)
				'***************************************************************
				
				'Marca: Q
				loRango = loCeldas(lnFilaActual, 17) 
				loRango.Value = CStr(loRenglon("Cod_Mar")).Trim()
				
				'Codigo 1: R
				loRango = loCeldas(lnFilaActual, 18) 
				loRango.Formula = "=LEFT(B" & lnFilaActual & ", $W$8)"
				
				'Codigo 2: S
				loRango = loCeldas(lnFilaActual, 19) 
				loRango.Formula = "=LEFT(B" & lnFilaActual & ", $X$8)"
				
				'Codigo 3: T
				loRango = loCeldas(lnFilaActual, 20) 
				loRango.Formula = "=LEFT(B" & lnFilaActual & ", $Y$8)"
				
				'Codigo 4: U
				loRango = loCeldas(lnFilaActual, 21) 
				loRango.Formula = "=LEFT(B" & lnFilaActual & ", $Z$8)"
				
				'Datos: V
				loRango = loCeldas(lnFilaActual, 22) 
				loRango.Formula = "=IF(SUM(D" & lnFilaActual & ":E" & lnFilaActual & ")+SUM(G" & lnFilaActual & ":H" & lnFilaActual & ")+N" & lnFilaActual & ">0,""a"","""")"
				
				'Inventario Actual $: W
				loRango = loCeldas(lnFilaActual, 23) 
				loRango.Value = "=K" & lnFilaActual & "*G" & lnFilaActual
				
				'Inventario por Llegar $: X
				loRango = loCeldas(lnFilaActual, 24) 
				loRango.Formula = "=K" & lnFilaActual & "*H" & lnFilaActual
	
				'Débil (HUESO): Y
				loRango = loCeldas(lnFilaActual, 25) 
				loRango.Formula = "=IF(E" & lnFilaActual & ">0, IF( ($J$2-I" & lnFilaActual & ")>30, ""H"", """"), """")"
					
				'Inventario Vendido $: Z
				loRango = loCeldas(lnFilaActual, 26) 
				loRango.Formula = "=K" & lnFilaActual & "*E" & lnFilaActual
				
				'Hueso por Marca: AA
				loRango = loCeldas(lnFilaActual, 27) 
				loRango.Formula = "=CONCATENATE(Q" & lnFilaActual & ",Y" & lnFilaActual & ")"
				
			Next lnRenglon
			
			'Borde de los datos
			Dim lnTotal As Integer = loDatos.Rows.Count
			loRango = loHoja.Range("B" & (lnFilaInicio+1) & ":O" & (lnFilaInicio + lnTotal))
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

			loRango = loHoja.Range("Q" & (lnFilaInicio+1) & ":AA" & (lnFilaInicio + lnTotal))
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)

			'Filtro automático en el encabezado
			loHoja.AutoFilterMode = False
			loRango = loHoja.Range("B" & lnFilaInicio & ":AA" & lnFilaInicio)
			loRango.Select()
			loExcel.Selection.AutoFilter()
			
			'Formato condicional: Efectividad
			loRango = loHoja.Range("B" & (lnFilaInicio+1) & ":C" & (lnFilaInicio + lnTotal))
			loRango.FormatConditions.Delete()
			loRango.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Nothing, "=$F" & (lnFilaInicio) & "<$AA$2")
			loRango.FormatConditions(1).Font.Color = &H0000FF 'Rojo

			loRango.FormatConditions.Add(Excel.XlFormatConditionType.xlExpression, Nothing, "=$F" & (lnFilaInicio) & ">$AA$3")
			loRango.FormatConditions(2).Font.Color = &HFF0000 'Azul

			'Formato condicional: Ganancia
			loRango = loHoja.Range("M" & (lnFilaInicio+1) & ":M" & (lnFilaInicio + lnTotal))
			loRango.FormatConditions.Delete()
			loRango.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlLess, "=$AA$4")
			loRango.FormatConditions(1).Font.Color = &H0000FF 'Rojo
			loRango.FormatConditions(1).Font.Bold = True

			
			'Totales
			Dim lnDesde AS Integer = lnFilaInicio+1
			Dim lnHasta AS Integer = lnFilaInicio + lnTotal
			
			lnFilaInicio += lnTotal + 2
			loRango = loHoja.Range("B" & (lnFilaInicio))
			loRango.NumberFormat = "@"
			loRango.Value = "Total Artículos: " & lnTotal.ToString()
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

			loRango = loHoja.Range("C" & (lnFilaInicio))
			loRango.NumberFormat = "@"
			loRango.Value = "Totales Generales: "
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

			loRango = loHoja.Range("D" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoMontos
			loRango.Value = "=SUM(D" & lnDesde & ":D" & lnHasta	& ")"

			loRango = loHoja.Range("E" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoMontos
			loRango.Value = "=SUM(E" & lnDesde & ":E" & lnHasta	& ")"

			loRango = loHoja.Range("F" & (lnFilaInicio))
			loRango.NumberFormat = "0.00"
			loRango.Value = "=IF(D" & lnFilaInicio & ">0, E" & lnFilaInicio	& "*100/D" & lnFilaInicio & ", 0)"

			loRango = loHoja.Range("G" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoCantidad
			loRango.Value = "=SUM(G" & lnDesde & ":G" & lnHasta	& ")"

			loRango = loHoja.Range("H" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoCantidad
			loRango.Value = "=SUM(H" & lnDesde & ":H" & lnHasta	& ")"
            
            loRango = loHoja.Range("N" & (lnDesde) & ":O" & lnHasta)
			loRango.Font.ColorIndex = 7 'Fucsia
			loRango.Font.Bold = True

			'***************************
			' Adicionales en Encabezado
			'***************************
			
			'Parametros: Encabezado	1
            loRango = loHoja.Range("W1:AA1")
            loRango.MergeCells = True
            loRango.Value = "PARÁMETROS DE ANÁLISIS DE EFECTIVIDAD"
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
            loRango.Font.Bold = True
            loRango.Font.Color = &HFFFFFF 'Blanco
            loRango.Interior.Color = &H000000 'Negro
            			
			'Parametros: % de Efectividad baja, menor a (25):
            loRango = loHoja.Range("W2:Z2")
            loRango.MergeCells = True
            loRango.Value = "% DE EFECTIVIDAD BAJA, MENOR A (25):"
            loRango = loHoja.Range("AA2")
            loRango.NumberFormat = "0.00"
            loRango.Value = Me.lnEfectividadBaja 
			
			'Parametros: % de Efectividad alta, mayor a (60):
            loRango = loHoja.Range("W3:Z3")
            loRango.MergeCells = True
            loRango.Value = "% DE EFECTIVIDAD ALTA, MAYOR A (60):"
            loRango = loHoja.Range("AA3")
            loRango.NumberFormat = "0.00"
            loRango.Value = Me.lnEfectividadAlta
			
			'Parametros: % de Ganancia baja, menor a (8):
            loRango = loHoja.Range("W4:Z4")
            loRango.MergeCells = True
            loRango.Value = "% DE GANANCIA BAJA, MENOR A (8):"
            loRango = loHoja.Range("AA4")
            loRango.NumberFormat = "0.00"
            loRango.Value = Me.lnGananciaBaja
			
			'Parametros: % Adicional de Compra sugerida (20):
            loRango = loHoja.Range("W5:Z5")
            loRango.MergeCells = True
            loRango.Value = "% ADICIONAL DE COMPRA SUGERIDA (20):"
            loRango = loHoja.Range("AA5")
            loRango.NumberFormat = "0.00"
            loRango.Value = Me.lnCompraAdicionalSugerida
			
			'Parametros: Días para Evaluar Inventario Débil (30):
            loRango = loHoja.Range("W6:Z6")
            loRango.MergeCells = True
            loRango.Value = "DÍAS PARA EVALUAR INV. DÉBIL (30):"
            loRango = loHoja.Range("AA6")
            loRango.NumberFormat = "0"
            loRango.Value = Me.lnDiasEvaluarInventario

			'Parametros: Encabezado 2
            loRango = loHoja.Range("W7:Z7")
            loRango.MergeCells = True
            loRango.Value = "LONGITUD DE COD 0, COD 1, COD2 Y COD 3"
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
            loRango.Font.Bold = True
            loRango.Font.Color = &HFFFFFF 'Blanco
            loRango.Interior.Color = &H000000 'Negro

            loRango = loHoja.Range("W8")
            loRango.NumberFormat = "0"
            loRango.Value = KN_LONGITUD_CODIGO_0

            loRango = loHoja.Range("X8")
            loRango.NumberFormat = "0"
            loRango.Value = KN_LONGITUD_CODIGO_1

            loRango = loHoja.Range("Y8")
            loRango.NumberFormat = "0"
            loRango.Value = KN_LONGITUD_CODIGO_2

            loRango = loHoja.Range("Z8")
            loRango.NumberFormat = "0"
            loRango.Value = KN_LONGITUD_CODIGO_3
			
			'Formato de Parámetros
            loRango = loHoja.Range("W2:AA6")
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
            loFuente = loRango.Font
            loFuente.Bold = True
			
            loRango = loHoja.Range("W8:Z8")
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
            With loRango.Borders(Excel.XlBordersIndex.xlInsideVertical)
				 .LineStyle = Excel.XlLineStyle.xlContinuous
				 .Weight = Excel.XlBorderWeight.xlThin 
			End With
            
			' Monto a Comprar
            loRango = loHoja.Range("N8:O8")
            loRango.MergeCells = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.Value = "MONTO POR COMPRAR"
            loRango.Interior.ColorIndex = 7 'Fucsia
            loFuente = loRango.Font
            loFuente.Size = 8
            loFuente.Color = &HFFFFFF 'Blanco

            loRango = loHoja.Range("N9:O9")
            loRango.MergeCells = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
			loRango.NumberFormat = lcFormatoMontos
            loRango.Value = "=SUM(O" & lnDesde & ":O" & lnHasta	& ")"
            loRango.Interior.ColorIndex = 7 'Fucsia
            loFuente = loRango.Font
            loFuente.Size = 8
            loFuente.Color = &HFFFFFF 'Blanco
			
			' Cotizado y Facturado (Encabezado)
            loRango = loHoja.Range("D1")
            loRango.Value = "COT"
            loRango.Interior.Color = &H000000 'Negro
            loFuente = loRango.Font
            loFuente.Color = &HFFFFFF 'Blanco
            
            loRango = loHoja.Range("E1")
            loRango.Value = "VEN"
            loRango.Interior.Color = &H000000 'Negro
            loFuente = loRango.Font
            loFuente.Color = &HFFFFFF 'Blanco

            loRango = loHoja.Range("D2")
			loRango.NumberFormat = lcFormatoCantidad
            loRango.Value = "=SUM(D" & lnDesde & ":D" & lnHasta	& ")"
            
            loRango = loHoja.Range("E2")
			loRango.NumberFormat = lcFormatoCantidad
            loRango.Value = "=SUM(E" & lnDesde & ":E" & lnHasta	& ")"
            
            loRango = loHoja.Range("D3:E3")
            loRango.MergeCells = True
            loRango.NumberFormat = "0.00%"
            loRango.Value = "=IF(D2>0,E2/D2,0)"

            loRango = loHoja.Range("D1:E3")
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
    		loRango.Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.XlLineStyle.xlLineStyleNone
    		loRango.Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.XlLineStyle.xlLineStyleNone
    
    		With loRango.Borders(Excel.XlBordersIndex.xlEdgeLeft)
    		    .LineStyle = Excel.XlLineStyle.xlContinuous
    		    .Weight = Excel.XlBorderWeight.xlMedium
    		    .Color = &H000000 'Negro
    		End With
    		With loRango.Borders(Excel.XlBordersIndex.xlEdgeTop)
    		    .LineStyle = Excel.XlLineStyle.xlContinuous
    		    .Weight = Excel.XlBorderWeight.xlMedium
    		    .Color = &H000000 'Negro
    		End With
    		With loRango.Borders(Excel.XlBordersIndex.xlEdgeBottom)
    		    .LineStyle = Excel.XlLineStyle.xlContinuous
    		    .Weight = Excel.XlBorderWeight.xlMedium
    		    .Color = &H000000 'Negro
    		End With
    		With loRango.Borders(Excel.XlBordersIndex.xlEdgeRight)
    		    .LineStyle = Excel.XlLineStyle.xlContinuous
    		    .Weight = Excel.XlBorderWeight.xlMedium
    		    .Color = &H000000 'Negro
    		End With
    		With loRango.Borders(Excel.XlBordersIndex.xlInsideVertical)
    		    .LineStyle = Excel.XlLineStyle.xlContinuous
    		    .Weight = Excel.XlBorderWeight.xlThin
    		    .Color = &H000000 'Negro
    		End With
    		With loRango.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
    		    .LineStyle = Excel.XlLineStyle.xlContinuous
    		    .Weight = Excel.XlBorderWeight.xlThin
    		    .Color = &H000000 'Negro
    		End With

			' Bloque de Resumen 1:
            loRango = loHoja.Range("G1:I1")
            loRango.MergeCells = True
            loRango.Value = "DÓLAR"
            
            loRango = loHoja.Range("J1:K1")
            loRango.MergeCells = True
            loRango.Interior.ColorIndex = 38
            
            loRango = loHoja.Range("G2:I2")
            loRango.MergeCells = True
            loRango.Value = "FECHA CORTE:"
            
            loRango = loHoja.Range("J2:K2")
            loRango.MergeCells = True
			loRango.NumberFormat = "[$-200A]dd-mmm-yyyy;@"		'--> [$-200A]= Español; [$-409A]= Inglés, etc...
			loRango.Value = CDate(ldFechaInicio.AddDays(-1)).ToString("dd-MMM-yyyy", loFormatoFecha)
            loRango.Interior.ColorIndex = 38
            
            loRango = loHoja.Range("G3:I3")
            loRango.MergeCells = True
            loRango.Value = "COSTO FACTURADO ESTIMADO:"
            
            loRango = loHoja.Range("J3:K3")
            loRango.MergeCells = True
			loRango.NumberFormat = lcFormatoMontos
			loRango.Value = "=SUMPRODUCT(E" & lnDesde & ":E" & lnHasta	& ",K" & lnDesde & ":K" & lnHasta	& ")"
            
            loRango = loHoja.Range("G4:I4")
            loRango.MergeCells = True
            loRango.Value = "DÍAS DE INVENTARIO CON TODO:"
            
            loRango = loHoja.Range("J4:K4")
            loRango.MergeCells = True
			loRango.NumberFormat = lcFormatoMontos
			loRango.Value = "=J3*$AA$6/O3"
            
            loRango = loHoja.Range("G5:I5")
            loRango.MergeCells = True
            loRango.Value = "DÍAS DE INVENTARIO FUERTE:"
            
            loRango = loHoja.Range("J5:K5")
            loRango.MergeCells = True
			loRango.NumberFormat = lcFormatoMontos
			loRango.Value = "=J3*$AA$6/(O3-O4)"
            
            loRango = loHoja.Range("G1:K2")
            loRango.Font.Bold = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
            
            loRango = loHoja.Range("G3:K5")
            loRango.Font.Bold = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
            

 			' Bloque de Resumen 2:
            loRango = loHoja.Range("M1:N1")
            loRango.MergeCells = True
            loRango.Value = "TOTAL COMPRA:"
            
            loRango = loHoja.Range("O1:P1")
            loRango.MergeCells = True
			loRango.NumberFormat = lcFormatoMontos
            loRango.Value = "=SUM(O" & lnDesde & ":O" & lnHasta	& ")"
            
            loRango = loHoja.Range("M2:N2")
            loRango.MergeCells = True
            loRango.Value = "INVENTARIO POR LLEGAR:"
            
            loRango = loHoja.Range("O2:P2")
            loRango.MergeCells = True
			loRango.NumberFormat = lcFormatoMontos
			loRango.Value = "=SUMPRODUCT(H" & lnDesde & ":H" & lnHasta	& ",K" & lnDesde & ":K" & lnHasta	& ")"
            
            loRango = loHoja.Range("M3:N3")
            loRango.MergeCells = True
            loRango.Value = "INVENTARIO EN ALMACÉN:"
            
            loRango = loHoja.Range("O3:P3")
            loRango.MergeCells = True
			loRango.NumberFormat = lcFormatoMontos
			loRango.Value = "=SUMPRODUCT(G" & lnDesde & ":G" & lnHasta	& ",K" & lnDesde & ":K" & lnHasta	& ")"
 
            loRango = loHoja.Range("M4:N4")
            loRango.MergeCells = True
            loRango.Value = "TOTAL INVENTARIO DÉBIL:"
            
            loRango = loHoja.Range("O4:P4")
            loRango.MergeCells = True
			loRango.NumberFormat = lcFormatoMontos
			loRango.Value = "=SUMIF(Y" & lnDesde & ":Y" & lnHasta & ",""H"", W" & lnDesde & ":W" & lnHasta & ")"
 
            loRango = loHoja.Range("M5:N5")
            loRango.MergeCells = True
            loRango.Value = "CANTIDAD PROD. DÉBILES:"
            
            loRango = loHoja.Range("O5")
			loRango.NumberFormat = lcFormatoCantidad
			loRango.Value = "=SUMIF(Y" & lnDesde & ":Y" & lnHasta & ",""H"", G" & lnDesde & ":G" & lnHasta & ")"
																			 
            loRango = loHoja.Range("P5")
			loRango.NumberFormat = "0.00%"
			loRango.Value = "=IF(O3>0,P4/O3,1)"
 
            loRango = loHoja.Range("M6:N6")
            loRango.MergeCells = True
            loRango.Value = "CANTIDAD ÍTEMS DÉBILES:"
            
            loRango = loHoja.Range("O6")
			loRango.NumberFormat = "0"
			loRango.Value = "=COUNTIF(Y" & lnDesde & ":Y" & lnHasta & ",""H"")"
			
            loRango = loHoja.Range("M7:N7")
            loRango.MergeCells = True
            loRango.Value = "DÉBILES VENDIDOS:"
            
            loRango = loHoja.Range("O7:P7")
            loRango.MergeCells = True
			loRango.NumberFormat = lcFormatoMontos
			loRango.Value = "=SUMIF(Y" & lnDesde & ":Y" & lnHasta & ",""H"", Z" & lnDesde & ":Z" & lnHasta & ")"
			
			
            loRango = loHoja.Range("M1:P3")
            loRango.Font.Bold = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
                        
            loRango = loHoja.Range("M4:P7")
            loRango.Font.Bold = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
            
            loRango = loHoja.Range("M6:P6")
            With loRango.Borders(Excel.XlBordersIndex.xlEdgeBottom)
				.Weight = Excel.XlBorderWeight.xlThin
				.LineStyle = Excel.XlLineStyle.xlDot
			End With
                        
            loRango = loHoja.Range("O5:O6")
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
			
            loRango = loHoja.Range("O4:P6")
            loRango.Interior.Color = &H00FFFF 'Amarillo
                        
            loRango = loHoja.Range("O7:P7")
            loRango.Interior.ColorIndex = 40 'Gris claro
           
 			' Bloque de Resumen 3:
            loRango = loHoja.Range("R1:U1")
            loRango.MergeCells = True
            loRango.Value = "TOTAL UNIDADES:"
            
            loRango = loHoja.Range("R2:U2")
            loRango.MergeCells = True
			loRango.NumberFormat = lcFormatoCantidad
            loRango.Formula = "=SUM(G" & lnDesde & ":G" & lnHasta	& ")"
            
            loRango = loHoja.Range("R3:U3")
            loRango.MergeCells = True
            loRango.Value = "TOTAL ITEMS CON STOCK"
            
            loRango = loHoja.Range("R4:U4")
            loRango.MergeCells = True
            loRango.Formula = "=COUNTIF(G" & lnDesde & ":G" & lnHasta	& ","">0"")"

            loRango = loHoja.Range("R5:U5")
            loRango.MergeCells = True
            loRango.Value = "INV + TRÁNSITO"
            
            loRango = loHoja.Range("R6:U6")
            loRango.MergeCells = True
			loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=O2+O3"

            loRango = loHoja.Range("R7:U7")
            loRango.MergeCells = True
            loRango.Value = "PRODUCTOS DÉBILES VENDIDOS"
            
            loRango = loHoja.Range("R8:U8")
            loRango.MergeCells = True
			loRango.NumberFormat = lcFormatoCantidad
			loRango.Formula = "=SUMIF(Y" & lnDesde & ":Y" & lnHasta & ",""H"", E" & lnDesde & ":E" & lnHasta & ")"
			loRango.Interior.ColorIndex = 40 'Gris claro
 
            loRango = loHoja.Range("R1:U2")
            loRango.Font.Bold = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
                        
            loRango = loHoja.Range("R3:U4")
            loRango.Font.Bold = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
            
            loRango = loHoja.Range("R5:U6")
            loRango.Font.Bold = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
                        
            loRango = loHoja.Range("R7:U8")
            loRango.Font.Bold = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThin)
                        
			'Ajustes finales: Ancho de columnas, alto de filas, márgenes, etc...
            loHoja.Range("D11").Select()
			loExcel.ActiveWindow.DisplayGridlines = False
			loExcel.ActiveWindow.FreezePanes = True
			
			loFilas = loCeldas.Rows
			loFilas.AutoFit()
			
			loColumnas = loCeldas.Rows
			loColumnas.AutoFit()
			
			loRango = loHoja.Range("B1:B" & lnFilaInicio)
			loRango.ColumnWidth = 23
			
			loRango = loHoja.Range("C1:C" & lnFilaInicio)
			loRango.ColumnWidth = 72
			
			loRango = loHoja.Range("D1:D" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("E1:E" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("F1:F" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("G1:G" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("H1:H" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("I1:I" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("J1:J" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("K1:K" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("L1:L" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("M1:M" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("N1:N" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("O1:O" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("Q1:U" & lnFilaInicio)
			loRango.ColumnWidth = 7
			
			loRango = loHoja.Range("W1:AA" & lnFilaInicio)
			loRango.ColumnWidth = 9
			
            ' Seleccionamos la primera celda del libro
			loRango = loHoja.Range("A1")
            loRango.Select()

            'Guardamos los cambios del libro activo
            loLibro.SaveAs(lcNombreArchivo)
            
		 '******************************************************************'
		 ' IMPORTANTE: Forma correcta de liberar recursos!!!				'
		 '******************************************************************'
            ' Cerramos y liberamos recursos

        Catch loExcepcion As Exception
			
			Throw New Exception("No fue posible exportar los datos a excel. " & loExcepcion.Message, loExcepcion)
			
        Finally

			If (loFuente IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loFuente)
				loFuente = Nothing
			End If
			
			If (loFormas IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loFormas)
				loFormas = Nothing
			End If
			
			If (loRango IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loRango)
				loRango = Nothing
			End If
			
			If (loFilas IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loFilas)
				loFilas = Nothing
			End If
			
			If (loColumnas IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loColumnas)
				loColumnas = Nothing
			End If
			
			If (loCeldas IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loCeldas)
				loCeldas = Nothing
			End If
			
			If (loHoja IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loHoja)
				loHoja = Nothing
			End If
			
			If (loLibro IsNot Nothing) Then
				loLibro.Close(True)
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loLibro)
				loLibro = Nothing
			End If

			If (laLibros IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(laLibros)
				laLibros = Nothing
			End If
            
            loExcel.Quit()

			System.Runtime.InteropServices.Marshal.ReleaseComObject(loExcel)
            loExcel = Nothing 
            
            GC.Collect()
            GC.WaitForPendingFinalizers()
            
        End Try

    End Sub


'''' <summary>
'''' Cierre y liberacion de recursos de los objetos de la libreria Excel
'''' </summary>
'''' <param name="objeto"></param>
'''' <remarks></remarks>
'    Private Sub mLiberar(ByVal objeto As Object)
'        Try
'            System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto)
'            objeto = Nothing
'        Catch ex As Exception
'            objeto = Nothing
'        End Try
'    End Sub

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
' RJG: 24/08/12: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' RJG: 03/09/12: Ajuste para mostrar artículos con stock que no tengan movimientos.			'
'-------------------------------------------------------------------------------------------'
' RJG: 23/04/14: Optimización en código de envío a Excel: calculo "manual" de fórmulas;     '
'                formato de celdas "fuera" del bucle FOR.			                        '
'-------------------------------------------------------------------------------------------'
