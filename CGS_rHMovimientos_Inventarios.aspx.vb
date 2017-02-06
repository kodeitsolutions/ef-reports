'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rHMovimientos_Inventarios "
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rHMovimientos_Inventarios
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

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
            'Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            'Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            'Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            'Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            'Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            'Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            'Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            'Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            'Dim lcParametro10Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))

            'Dim lcParametro11Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))

            'Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            'Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            'Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            'Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))

            'Dim lcParametroFIDesde As String = goServicios.mObtenerCampoFormatoSQL(New Date (1990, 01, 01), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            'Dim lcParametroFIHasta As String = goServicios.mObtenerCampoFormatoSQL((CDATE(cusAplicacion.goReportes.paParametrosIniciales(1)).AddDays(-1)), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            'Dim lcCosto As String = "Cos_Pro1"
            Dim loComandoSeleccionar As New StringBuilder()

            'Select Case lcParametro11Desde
            '    Case "'Promedio MP'"
            '        lcCosto = "Cos_Pro1"
            '    Case "'Ultimo MP'"
            '        lcCosto = "Cos_Ult1"
            '    Case "'Anterior MP'"
            '        lcCosto = "Cos_Ant1"
            '    Case "'Promedio MS'"
            '        lcCosto = "Cos_Pro2"
            '    Case "'Ultimo MS'"
            '        lcCosto = "Cos_Ult2"
            '    Case "'Anterior MS'"
            '        lcCosto = "Cos_Ant2"
            'End Select

            ''-------------------------------------------------------------------------------------------'
            '' Invocacion del Store Procedure
            ''-------------------------------------------------------------------------------------------'
            'Dim lcAlmacenTraslado As String = goServicios.mObtenerCampoFormatoSQL(cusadministrativo.goAlmacen.pcAlmacenTransito)

            'loComandoSeleccionar.AppendLine("DECLARE @tmpArticulos AS TABLE(Cod_Art CHAR(30), Saldo DECIMAL(28,10)) ;")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("INSERT 	INTO @tmpArticulos ")
            'loComandoSeleccionar.AppendLine("SELECT		Cod_Art, 0")
            'loComandoSeleccionar.AppendLine("FROM		Articulos")
            'loComandoSeleccionar.AppendLine("WHERE		Cod_Art 	BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            'loComandoSeleccionar.AppendLine("		AND	Cod_Dep 	BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            'loComandoSeleccionar.AppendLine("		AND	Cod_Sec 	BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            'loComandoSeleccionar.AppendLine("		AND	Cod_Mar 	BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            'loComandoSeleccionar.AppendLine("		AND	Cod_Cla 	BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            'loComandoSeleccionar.AppendLine("		AND	Cod_Tip 	BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            'loComandoSeleccionar.AppendLine("		AND	Cod_Pro 	BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            'loComandoSeleccionar.AppendLine("		AND	Cod_Uni1	BETWEEN " & lcParametro12Desde & " AND " & lcParametro12Hasta)
            'loComandoSeleccionar.AppendLine("		AND	Cod_Ubi		BETWEEN " & lcParametro13Desde & " AND " & lcParametro13Hasta)
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")

            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")



            ''Me.mEscribirConsulta(loComandoSeleccionar.ToString())


            ''-------------------------------------------------------------------------------------------'
            '' El código del Almacén de Tránsito se requiere para los Traslados entre Almacénes          '
            '' Confirmados o Procesados.	                                                                '
            ''-------------------------------------------------------------------------------------------'
            'Dim lcAlmacenTransito As String
            'lcAlmacenTransito = goServicios.mObtenerCampoFormatoSQL(goOpciones.mObtener("CODALMTRA", "C"))

            'Select Case lcParametro10Desde

            '    Case "'Todos'"

            '        ' Select de la tabla de Ajustes
            '        loComandoSeleccionar.AppendLine("SELECT		'Ajustes'								AS	Operacion, 	")
            '        loComandoSeleccionar.AppendLine("			Ajustes.Documento						AS	Documento, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Art				AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine("			Ajustes.Fec_Ini							AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Renglon				AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine("			'No Aplica'								AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Alm				AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine("			(CASE WHEN Renglones_Ajustes.Tipo = 'Salida' THEN Renglones_Ajustes.Can_Art1 ELSE 0.0 END)		AS	Can_Sal, ")
            '        loComandoSeleccionar.AppendLine("			(CASE WHEN Renglones_Ajustes.Tipo = 'Entrada'  THEN Renglones_Ajustes.Can_Art1 ELSE 0.0 END)		AS	Can_Ent, ")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Tipo					AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Ajustes." & lcCosto & "       AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	            AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Ajustes.Cod_Suc							AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            '        loComandoSeleccionar.AppendLine("FROM		Ajustes")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Ajustes.Status					=	'Confirmado' ")
            '        loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Tipo			IN	('Entrada', 'Salida') ")
            '        loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Fec_Ini					<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Cod_Alm		BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Cod_Suc					BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro9Hasta)

            '        ' Union con Select de la tabla de Traslados en las Salidas
            '        loComandoSeleccionar.AppendLine("UNION ALL ")
            '        loComandoSeleccionar.AppendLine("SELECT		'Traslados'								AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Documento						AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cod_Art				AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Fec_Ini						AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Renglon				AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			'No Aplica'								AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Alm_Ori						AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Can_Art1			AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Salida'								AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados." & lcCosto & "     AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	            AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Cod_Suc						AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("FROM		Traslados")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Traslados.Status		IN ('Confirmado', 'Procesado')	")
            '        loComandoSeleccionar.AppendLine(" 		AND Traslados.Fec_Ini				<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Traslados.Alm_Ori		BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Traslados.Cod_Suc				BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

            '        ' Union con Select de la tabla de Traslados en las Entradas
            '        loComandoSeleccionar.AppendLine("UNION ALL ")
            '        loComandoSeleccionar.AppendLine("SELECT		'Traslados'								AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Documento						AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cod_Art				AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Fec_Ini						AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Renglon				AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			'No Aplica'								AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			CASE Traslados.Status					")
            '        loComandoSeleccionar.AppendLine(" 			    WHEN 'Confirmado'	THEN " & lcAlmacenTransito)
            '        loComandoSeleccionar.AppendLine(" 			    WHEN 'Procesado'	THEN Traslados.Alm_Des")
            '        loComandoSeleccionar.AppendLine(" 			END										AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Can_Art1			AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Entrada'								AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados." & lcCosto & "     AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	            AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Cod_Suc						AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("FROM		Traslados")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Traslados.Status			IN ('Confirmado', 'Procesado')	")
            '        loComandoSeleccionar.AppendLine(" 		AND Traslados.Fec_Ini			<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND (CASE Traslados.Status					")
            '        loComandoSeleccionar.AppendLine(" 			    WHEN 'Confirmado'	THEN " & lcAlmacenTransito)
            '        loComandoSeleccionar.AppendLine(" 			    WHEN 'Procesado'	THEN Traslados.Alm_Des")
            '        loComandoSeleccionar.AppendLine(" 			END)						BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Traslados.Cod_Suc			BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

            '        ' Union con Select de la tabla de Entregas
            '        loComandoSeleccionar.AppendLine("UNION ALL ")
            '        loComandoSeleccionar.AppendLine("SELECT		'Entregas'							AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Entregas.Documento					AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cod_Art			AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Entregas.Fec_Ini					AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Renglon			AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			Entregas.Cod_Cli					AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cod_Alm			AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Can_Art1			AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Salida'							AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Entregas." & lcCosto & "  AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	        AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Entregas.Cod_Suc					AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("FROM		Entregas")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Entregas ON Renglones_Entregas.Documento = Entregas.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Entregas.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Entregas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine(" 		AND Entregas.Fec_Ini				<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Renglones_Entregas.Cod_Alm		BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Entregas.Cod_Suc				BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

            '        ' Union con Select de la tabla de Facturas
            '        loComandoSeleccionar.AppendLine("UNION ALL ")
            '        loComandoSeleccionar.AppendLine("SELECT		'Facturas'							AS	Operacion, ")
            '        loComandoSeleccionar.AppendLine(" 			Facturas.Documento					AS	Documento, ")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cod_Art			AS	Cod_Art, ")
            '        loComandoSeleccionar.AppendLine(" 			Facturas.Fec_Ini					AS	Fec_Ini, ")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Renglon			AS	Renglon, ")
            '        loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Cli					AS	Cliente, ")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cod_Alm			AS	Cod_Alm, ")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Can_Art1			AS	Can_Sal, ")
            '        loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Ent, ")
            '        loComandoSeleccionar.AppendLine(" 			'Salida'							AS	Tipo, ")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Facturas." & lcCosto & "  AS	Cos_Pro, ")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	        AS	Costo, ")
            '        loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Suc					AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo ")
            '        loComandoSeleccionar.AppendLine("FROM		Facturas")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Facturas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine("		AND Renglones_Facturas.Tip_Ori		<>	'Entregas' ")
            '        loComandoSeleccionar.AppendLine("		AND Facturas.Fec_Ini				<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Renglones_Facturas.Cod_Alm		BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Suc				BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

            '        ' Union con Select de la tabla de Recepciones
            '        loComandoSeleccionar.AppendLine("UNION ALL ")
            '        loComandoSeleccionar.AppendLine("SELECT		'Recepciones'						AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Recepciones.Documento				AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cod_Art		AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Recepciones.Fec_Ini					AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Renglon		AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			Recepciones.Cod_Pro					AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cod_Alm		AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Can_Art1		AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Entrada'							AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones." & lcCosto & " AS Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	        AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Recepciones.Cod_Suc					AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo ")
            '        loComandoSeleccionar.AppendLine("FROM		Recepciones")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Recepciones ON Renglones_Recepciones.Documento = Recepciones.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Recepciones.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Recepciones.Status			IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine("		AND Recepciones.Fec_Ini				<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Renglones_Recepciones.Cod_Alm	BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Recepciones.Cod_Suc				BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

            '        ' Union con Select de la tabla de Compras
            '        loComandoSeleccionar.AppendLine("UNION ALL ")
            '        loComandoSeleccionar.AppendLine("SELECT		'Compras'							AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine("			Compras.Documento					AS	Documento,	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Art			AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine("			Compras.Fec_Ini						AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Compras.Renglon			AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine("			Compras.Cod_Pro						AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Alm			AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine("			0.0									AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Compras.Can_Art1			AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine("			'Entrada'							AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine("			Renglones_Compras." & lcCosto & "   AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine("			" & lcParametro11Desde & "	        AS	Costo,		")
            '        loComandoSeleccionar.AppendLine("			Compras.Cod_Suc						AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("FROM		Compras")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Compras.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Compras.Status				IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine(" 		AND Renglones_Compras.Tip_Ori	<>	'Recepciones' ")
            '        loComandoSeleccionar.AppendLine("		AND Compras.Fec_Ini				<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Renglones_Compras.Cod_Alm	BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Compras.Cod_Suc				BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

            '        ' Union con Select de la tabla de Devoluciones_Clientes
            '        loComandoSeleccionar.AppendLine(" UNION ALL ")
            '        loComandoSeleccionar.AppendLine(" SELECT	'Dev_Cli'								AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Documento			AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cod_Art				AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Fec_Ini			AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Renglon				AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Cod_Cli			AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cod_Alm				AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Can_Art1			AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Entrada'								AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DClientes." & lcCosto & "     AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	            AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Cod_Suc			AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("FROM		Devoluciones_Clientes")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_DClientes ON Renglones_DClientes.Documento = Devoluciones_Clientes.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_DClientes.Cod_Art ")
            '        loComandoSeleccionar.AppendLine(" WHERE			Devoluciones_Clientes.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine(" 			AND Devoluciones_Clientes.Fec_Ini		<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 			AND Renglones_DClientes.Cod_Alm			BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 			AND Devoluciones_Clientes.Cod_Suc		BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 				AND " & lcParametro9Hasta)

            '        ' Union con Select de la tabla de Devoluciones_Proveedores
            '        loComandoSeleccionar.AppendLine("UNION ALL ")
            '        loComandoSeleccionar.AppendLine("SELECT		'Dev_Pro'								AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Documento		AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cod_Art			AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Fec_Ini		AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Renglon			AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Cod_Pro		AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cod_Alm			AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Can_Art1			AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Salida'								AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores." & lcCosto & "  AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	            AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Cod_Suc        AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("FROM		Devoluciones_Proveedores")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_DProveedores ON Renglones_DProveedores.Documento = Devoluciones_Proveedores.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_DProveedores.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Devoluciones_Proveedores.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Proveedores.Fec_Ini	<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Renglones_DProveedores.Cod_Alm		BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Proveedores.Cod_Suc					BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

            '    Case "'Ajustes_Entrada'"

            '        ' Select de la tabla de Ajustes solo para las Entradas
            '        loComandoSeleccionar.AppendLine("SELECT		'Ajustes'								AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine("			Ajustes.Documento						AS	Documento,	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Art				AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine("			Ajustes.Fec_Ini							AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Renglon				AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine("			'No Aplica'								AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Alm				AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine("			0.0										AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Can_Art1				AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Tipo					AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Ajustes." & lcCosto & "       AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	            AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Ajustes.Cod_Suc							AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            '        loComandoSeleccionar.AppendLine("FROM		Ajustes")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Ajustes.Status					=	'Confirmado' ")
            '        loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Tipo			=	'Entrada' ")
            '        loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Fec_Ini					<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Cod_Alm		BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Cod_Suc					BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro9Hasta)

            '    Case "'Ajustes_Salida'"

            '        ' Select de la tabla de Ajustes solo para las Salidas
            '        loComandoSeleccionar.AppendLine("SELECT		'Ajustes'								AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine("			Ajustes.Documento						AS	Documento,	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Art				AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine("			Ajustes.Fec_Ini							AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Renglon				AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine("			'No Aplica'								AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Alm				AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Can_Art1				AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine("			0.0										AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Tipo					AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Ajustes." & lcCosto & "       AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	            AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Ajustes.Cod_Suc							AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            '        loComandoSeleccionar.AppendLine("FROM		Ajustes")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Ajustes.Status					=	'Confirmado' ")
            '        loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Tipo			=	'Salida' ")
            '        loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Fec_Ini					<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Cod_Alm		BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Cod_Suc					BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro9Hasta)

            '    Case "'Traslados_Salida'"

            '        ' Union con Select de la tabla de Traslados en las Salidas
            '        loComandoSeleccionar.AppendLine("SELECT		'Traslados'								AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Documento						AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cod_Art				AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Fec_Ini						AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Renglon				AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			'No Aplica'								AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Alm_Ori						AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Can_Art1			AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Salida'								AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados." & lcCosto & "     AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	            AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Cod_Suc						AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            '        loComandoSeleccionar.AppendLine("FROM		Traslados")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Traslados.Status		IN ('Confirmado', 'Procesado')	")
            '        loComandoSeleccionar.AppendLine(" 		AND Traslados.Fec_Ini				<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Traslados.Alm_Ori		BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Traslados.Cod_Suc				BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

            '    Case "'Traslados_Entrada'"

            '        ' Union con Select de la tabla de Traslados en las Entradas
            '        loComandoSeleccionar.AppendLine("SELECT		'Traslados'								AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Documento						AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cod_Art				AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Fec_Ini						AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Renglon				AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			'No Aplica'								AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			CASE Traslados.Status					")
            '        loComandoSeleccionar.AppendLine(" 			    WHEN 'Confirmado'	THEN " & lcAlmacenTransito)
            '        loComandoSeleccionar.AppendLine(" 			    WHEN 'Procesado'	THEN Traslados.Alm_Des")
            '        loComandoSeleccionar.AppendLine(" 			END										AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Can_Art1			AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Entrada'								AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Traslados." & lcCosto & "     AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	            AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Traslados.Cod_Suc						AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            '        loComandoSeleccionar.AppendLine("FROM		Traslados")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Traslados.Status			IN ('Confirmado', 'Procesado')	")
            '        loComandoSeleccionar.AppendLine(" 		AND Traslados.Fec_Ini			<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND (CASE Traslados.Status					")
            '        loComandoSeleccionar.AppendLine(" 			    WHEN 'Confirmado'	THEN " & lcAlmacenTransito)
            '        loComandoSeleccionar.AppendLine(" 			    WHEN 'Procesado'	THEN Traslados.Alm_Des")
            '        loComandoSeleccionar.AppendLine(" 			END)						BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Traslados.Cod_Suc			BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

            '    Case "'Facturas_Venta'"

            '        ' Union con Select de la tabla de Facturas
            '        loComandoSeleccionar.AppendLine("SELECT		'Facturas'							AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Facturas.Documento					AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cod_Art			AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Facturas.Fec_Ini					AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Renglon			AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Cli					AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cod_Alm			AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Can_Art1			AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Salida'							AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Facturas." & lcCosto & "  AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	        AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Suc					AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            '        loComandoSeleccionar.AppendLine("FROM		Facturas")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Facturas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine("		AND Renglones_Facturas.Tip_Ori		<>	'Entregas' ")
            '        loComandoSeleccionar.AppendLine("		AND Facturas.Fec_Ini				<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Renglones_Facturas.Cod_Alm		BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Suc				BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

            '    Case "'Facturas_Compra'"

            '        ' Union con Select de la tabla de Compras
            '        loComandoSeleccionar.AppendLine("SELECT		'Compras'							AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine("			Compras.Documento					AS	Documento,	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Art			AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine("			Compras.Fec_Ini						AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Compras.Renglon			AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine("			Compras.Cod_Pro						AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Alm			AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine("			0.0									AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine("			Renglones_Compras.Can_Art1			AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine("			'Entrada'							AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine("			Renglones_Compras." & lcCosto & "   AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine("			" & lcParametro11Desde & "	        AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Compras.Cod_Suc						AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo	")
            '        loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            '        loComandoSeleccionar.AppendLine("FROM		Compras")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Compras.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Compras.Status				IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine(" 		AND Renglones_Compras.Tip_Ori	<>	'Recepciones' ")
            '        loComandoSeleccionar.AppendLine("		AND Compras.Fec_Ini				<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Renglones_Compras.Cod_Alm	BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Compras.Cod_Suc				BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

            '    Case "'Notas_Entrega'"

            '        ' Union con Select de la tabla de Entregas
            '        loComandoSeleccionar.AppendLine("SELECT		'Entregas'							AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Entregas.Documento					AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cod_Art			AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Entregas.Fec_Ini					AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Renglon			AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			Entregas.Cod_Cli					AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cod_Alm			AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Can_Art1			AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Salida'							AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Entregas." & lcCosto & "  AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	        AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Entregas.Cod_Suc					AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            '        loComandoSeleccionar.AppendLine("FROM		Entregas")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Entregas ON Renglones_Entregas.Documento = Entregas.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Entregas.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Entregas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine(" 		AND Entregas.Fec_Ini				<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Renglones_Entregas.Cod_Alm		BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Entregas.Cod_Suc				BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

            '    Case "'Notas_Recepcion'"

            '        ' Union con Select de la tabla de Recepciones
            '        loComandoSeleccionar.AppendLine("SELECT		'Recepciones'						AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Recepciones.Documento				AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cod_Art		AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Recepciones.Fec_Ini					AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Renglon		AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			Recepciones.Cod_Pro					AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cod_Alm		AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Can_Art1		AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Entrada'							AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones." & lcCosto & " AS Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	        AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Recepciones.Cod_Suc					AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo	")
            '        loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            '        loComandoSeleccionar.AppendLine("FROM		Recepciones")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_Recepciones ON Renglones_Recepciones.Documento = Recepciones.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Recepciones.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Recepciones.Status			IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine("		AND Recepciones.Fec_Ini				<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Renglones_Recepciones.Cod_Alm	BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine("		AND Recepciones.Cod_Suc				BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

            '    Case "'Devolucion_Ventas'"

            '        ' Union con Select de la tabla de Devoluciones_Clientes
            '        loComandoSeleccionar.AppendLine(" SELECT	'Dev_Cli'								AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Documento			AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cod_Art				AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Fec_Ini			AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Renglon				AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Cod_Cli			AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cod_Alm				AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Can_Art1			AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Entrada'								AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DClientes." & lcCosto & "     AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	            AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Cod_Suc			AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            '        loComandoSeleccionar.AppendLine("FROM		Devoluciones_Clientes")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_DClientes ON Renglones_DClientes.Documento = Devoluciones_Clientes.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_DClientes.Cod_Art ")
            '        loComandoSeleccionar.AppendLine(" WHERE			Devoluciones_Clientes.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine(" 			AND Devoluciones_Clientes.Fec_Ini		<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 			AND Renglones_DClientes.Cod_Alm			BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 			AND Devoluciones_Clientes.Cod_Suc		BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 				AND " & lcParametro9Hasta)

            '    Case "'Devolucion_Compras'"

            '        ' Union con Select de la tabla de Devoluciones_Proveedores
            '        loComandoSeleccionar.AppendLine("SELECT		'Dev_Pro'								AS	Operacion,	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Documento		AS	Documento,	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cod_Art			AS	Cod_Art, 	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Fec_Ini		AS	Fec_Ini, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Renglon			AS	Renglon, 	")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Cod_Pro		AS	Cliente, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cod_Alm			AS	Cod_Alm, 	")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Can_Art1			AS	Can_Sal, 	")
            '        loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Ent, 	")
            '        loComandoSeleccionar.AppendLine(" 			'Salida'								AS	Tipo,		")
            '        loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores." & lcCosto & "  AS	Cos_Pro,	")
            '        loComandoSeleccionar.AppendLine(" 			" & lcParametro11Desde & "	            AS	Costo,		")
            '        loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Cod_Suc		AS	Cod_Suc,	")
            '        loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
            '        loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            '        loComandoSeleccionar.AppendLine("FROM		Devoluciones_Proveedores")
            '        loComandoSeleccionar.AppendLine("	JOIN	Renglones_DProveedores ON Renglones_DProveedores.Documento = Devoluciones_Proveedores.Documento")
            '        loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_DProveedores.Cod_Art ")
            '        loComandoSeleccionar.AppendLine("WHERE		Devoluciones_Proveedores.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
            '        loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Proveedores.Fec_Ini	<=	" & lcParametro1Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Renglones_DProveedores.Cod_Alm		BETWEEN " & lcParametro2Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            '        loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Proveedores.Cod_Suc	BETWEEN " & lcParametro9Desde)
            '        loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

            'End Select

            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("-- Crea un índice para acelerar las siguientes operaciones")
            'loComandoSeleccionar.AppendLine("CREATE CLUSTERED INDEX PK_Fecha ON #curTemporal(Fec_Ini)")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("-- Calcula los saldos iniciales")
            'loComandoSeleccionar.AppendLine("UPDATE		#curTemporal")
            'loComandoSeleccionar.AppendLine("SET			Saldo = S.Saldo")
            'loComandoSeleccionar.AppendLine("FROM	(	SELECT	Cod_Art, Cod_Alm, Cod_Suc,")
            'loComandoSeleccionar.AppendLine("					SUM(Can_Ent - Can_Sal) As Saldo		")
            'loComandoSeleccionar.AppendLine("			FROM	#curTemporal")
            'loComandoSeleccionar.AppendLine("			WHERE	#curTemporal.Fec_Ini < " & lcParametro1Desde)
            'loComandoSeleccionar.AppendLine("			GROUP BY Cod_Art, Cod_Alm, Cod_Suc")
            'loComandoSeleccionar.AppendLine("		)	AS S")
            'loComandoSeleccionar.AppendLine("WHERE 		#curTemporal.Fec_Ini >= " & lcParametro1Desde)
            'loComandoSeleccionar.AppendLine("	AND		#curTemporal.Cod_Art = S.Cod_Art  ")
            'loComandoSeleccionar.AppendLine("	AND		#curTemporal.Cod_Alm = S.Cod_Alm ")
            'loComandoSeleccionar.AppendLine("	AND		#curTemporal.Cod_Suc = S.Cod_Suc")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("SELECT		#curTemporal.Saldo									AS Inicial,		")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Saldo									AS Saldo,		")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Operacion								AS Operacion,	")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Documento								AS Documento,	")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Cod_Art								AS Cod_Art, 	")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Fec_Ini								AS Fec_Ini, 	")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Renglon								AS Renglon, 	")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Cliente								AS Cliente, 	")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Cod_Alm								AS Cod_Alm, 	")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Can_Sal								AS Can_Sal, 	")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Can_Ent								AS Can_Ent, 	")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Tipo									AS Tipo,		")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Cos_Pro								AS Cos_Pro,		")
            'loComandoSeleccionar.AppendLine("			(#curTemporal.Cos_Pro * #curTemporal.Can_Sal)		AS Cos_Pro_Sal,")
            'loComandoSeleccionar.AppendLine("			(#curTemporal.Cos_Pro * #curTemporal.Can_Ent)		AS Cos_Pro_Ent,")
            'loComandoSeleccionar.AppendLine("			(	(#curTemporal.Cos_Pro * #curTemporal.Can_Ent)					")
            'loComandoSeleccionar.AppendLine("			  - (#curTemporal.Cos_Pro * #curTemporal.Can_Sal))	AS	Cos_Tot,	")
            'loComandoSeleccionar.AppendLine("			Articulos.Nom_Art									AS Nom_Art,		")
            'loComandoSeleccionar.AppendLine("			#curTemporal.Costo									AS Costo		")
            'loComandoSeleccionar.AppendLine("FROM		#curTemporal")
            'loComandoSeleccionar.AppendLine("	JOIN	Articulos ")
            'loComandoSeleccionar.AppendLine("		ON	Articulos.Cod_Art = #curTemporal.Cod_Art ")
            'loComandoSeleccionar.AppendLine("WHERE 		#curTemporal.Fec_Ini >= " & lcParametro1Desde)
            'loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("DROP TABLE #curTemporal")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("")


            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Desde AS VARCHAR(8) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodArt_Hasta AS VARCHAR(8) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Desde AS VARCHAR(10) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodAlm_Hasta AS VARCHAR(10) = " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Desde AS VARCHAR(2) = " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodDep_Hasta AS VARCHAR(2) = " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Desde AS VARCHAR(2) = " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodSec_Hasta AS VARCHAR(2) = " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodLot_Desde AS VARCHAR(30) = " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodLot_Hasta AS VARCHAR(30) = " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @tmpArticulos AS TABLE(Cod_Art CHAR(30), Saldo DECIMAL(28,10)) ;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO @tmpArticulos ")
            loComandoSeleccionar.AppendLine("SELECT	Cod_Art, 0")
            loComandoSeleccionar.AppendLine("FROM Articulos")
            loComandoSeleccionar.AppendLine("WHERE Cod_Art BETWEEN @lcCodArt_Desde AND @lcCodArt_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cod_Dep BETWEEN @lcCodDep_Desde AND @lcCodDep_Hasta")
            loComandoSeleccionar.AppendLine("	AND	Cod_Sec BETWEEN @lcCodSec_Desde AND @lcCodSec_Hasta")
            loComandoSeleccionar.AppendLine("	--AND Cod_Pro BETWEEN '' AND 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("		")
            loComandoSeleccionar.AppendLine("SELECT 'Ajustes'								AS	Operacion, 	")
            loComandoSeleccionar.AppendLine("		Ajustes.Documento						AS	Documento,")
            loComandoSeleccionar.AppendLine("		Renglones_Ajustes.Cod_Art				AS	Cod_Art,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot				AS	Lote, 	")
            loComandoSeleccionar.AppendLine("		Ajustes.Fec_Ini							AS	Fec_Ini, 	")
            loComandoSeleccionar.AppendLine("		Renglones_Ajustes.Renglon				AS	Renglon,  	")
            loComandoSeleccionar.AppendLine("		Renglones_Ajustes.Cod_Alm				AS	Cod_Alm, 	")
            'loComandoSeleccionar.AppendLine("		(CASE WHEN Renglones_Ajustes.Tipo = 'Salida' THEN Renglones_Ajustes.Can_Art1 ELSE 0.0 END)		AS	CanRng_Sal, ")
            'loComandoSeleccionar.AppendLine("		(CASE WHEN Renglones_Ajustes.Tipo = 'Entrada'  THEN Renglones_Ajustes.Can_Art1 ELSE 0.0 END)		AS	CanRng_Ent, ")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Renglones_Ajustes.Tipo = 'Salida' THEN Operaciones_Lotes.Cantidad ELSE 0.0 END)		AS	CanLte_Sal, ")
            loComandoSeleccionar.AppendLine("		(CASE WHEN Renglones_Ajustes.Tipo = 'Entrada'  THEN Operaciones_Lotes.Cantidad ELSE 0.0 END)		AS	CanLte_Ent, ")
            loComandoSeleccionar.AppendLine("		Renglones_Ajustes.Tipo					AS	Tipo,		")
            loComandoSeleccionar.AppendLine(" 		Articulos.Saldo				            AS	Saldo		")
            loComandoSeleccionar.AppendLine("INTO #curTemporal ")
            loComandoSeleccionar.AppendLine("FROM Ajustes")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento")
            loComandoSeleccionar.AppendLine("	JOIN @tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art ")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Ajustes.Documento")
            loComandoSeleccionar.AppendLine("       AND Operaciones_Lotes.Tip_Doc = 'Ajustes_Inventarios'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = Renglones_Ajustes.Tipo")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Ajustes.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Renglones_Ajustes.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE Ajustes.Status = 'Confirmado' ")
            loComandoSeleccionar.AppendLine(" 	AND	Renglones_Ajustes.Tipo IN ('Entrada', 'Salida') ")
            loComandoSeleccionar.AppendLine(" 	AND	Ajustes.Fec_Ini <= @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine(" 	AND	Renglones_Ajustes.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("   AND Operaciones_Lotes.Cod_Lot BETWEEN @lcCodLot_Desde AND @lcCodLot_Hasta")
            loComandoSeleccionar.AppendLine(" 	")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("	")
            loComandoSeleccionar.AppendLine("SELECT 'Traslados'								AS	Operacion,	")
            loComandoSeleccionar.AppendLine(" 		Traslados.Documento						AS	Documento,	")
            loComandoSeleccionar.AppendLine(" 		Renglones_Traslados.Cod_Art				AS	Cod_Art,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot				AS	Lote, 	")
            loComandoSeleccionar.AppendLine(" 		Traslados.Fec_Ini						AS	Fec_Ini, 	")
            loComandoSeleccionar.AppendLine(" 		Renglones_Traslados.Renglon				AS	Renglon,  	")
            loComandoSeleccionar.AppendLine(" 		Traslados.Alm_Ori						AS	Cod_Alm, 	")
            'loComandoSeleccionar.AppendLine(" 		Renglones_Traslados.Can_Art1			AS	CanRng_Sal, ")
            'loComandoSeleccionar.AppendLine(" 		0.0										AS	CanRng_Ent, ")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cantidad				AS	CanLte_Sal, ")
            loComandoSeleccionar.AppendLine(" 		0.0										AS	CanLte_Ent,	")
            loComandoSeleccionar.AppendLine(" 		'Salida'								AS	Tipo,	")
            loComandoSeleccionar.AppendLine(" 		Articulos.Saldo				            AS	Saldo		")
            loComandoSeleccionar.AppendLine("FROM Traslados")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
            loComandoSeleccionar.AppendLine("	JOIN @tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Traslados.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Traslados'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Salida'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Traslados.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Renglones_Traslados.Cod_Art ")
            loComandoSeleccionar.AppendLine("WHERE Traslados.Status IN ('Confirmado', 'Procesado')	")
            loComandoSeleccionar.AppendLine(" 	AND Traslados.Fec_Ini <= @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine(" 	AND Traslados.Alm_Ori BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("   AND Operaciones_Lotes.Cod_Lot BETWEEN @lcCodLot_Desde AND @lcCodLot_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT 'Traslados'								AS	Operacion,	")
            loComandoSeleccionar.AppendLine(" 		Traslados.Documento						AS	Documento,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Traslados.Cod_Art				AS	Cod_Art,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot				AS	Lote,  	")
            loComandoSeleccionar.AppendLine(" 		Traslados.Fec_Ini						AS	Fec_Ini, 	")
            loComandoSeleccionar.AppendLine(" 		Renglones_Traslados.Renglon				AS	Renglon, 	 	")
            loComandoSeleccionar.AppendLine(" 		CASE Traslados.Status					")
            loComandoSeleccionar.AppendLine(" 		    WHEN 'Confirmado'	THEN 'TRANSITO'")
            loComandoSeleccionar.AppendLine(" 		    WHEN 'Procesado'	THEN Traslados.Alm_Des")
            loComandoSeleccionar.AppendLine(" 		END										AS	Cod_Alm, 	")
            'loComandoSeleccionar.AppendLine(" 		0.0										AS	CanRng_Sal, ")
            'loComandoSeleccionar.AppendLine(" 		Renglones_Traslados.Can_Art1			AS	CanRng_Ent,")
            loComandoSeleccionar.AppendLine("		0.0										AS	CanLte_Sal, 	")
            loComandoSeleccionar.AppendLine(" 		Operaciones_Lotes.Cantidad				AS	CanLte_Ent, 	")
            loComandoSeleccionar.AppendLine(" 		'Entrada'								AS	Tipo,	")
            loComandoSeleccionar.AppendLine(" 		Articulos.Saldo				            AS	Saldo		")
            loComandoSeleccionar.AppendLine("FROM   Traslados")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
            loComandoSeleccionar.AppendLine("	JOIN @tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art ")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Traslados.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Traslados'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Traslados.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Renglones_Traslados.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE Traslados.Status IN ('Confirmado', 'Procesado')	")
            loComandoSeleccionar.AppendLine("   AND Traslados.Fec_Ini <= @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine(" 	AND (CASE Traslados.Status					")
            loComandoSeleccionar.AppendLine(" 		    WHEN 'Confirmado' THEN 'TRANSITO'")
            loComandoSeleccionar.AppendLine(" 		    WHEN 'Procesado' THEN Traslados.Alm_Des")
            loComandoSeleccionar.AppendLine(" 		END) BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("   AND Operaciones_Lotes.Cod_Lot BETWEEN @lcCodLot_Desde AND @lcCodLot_Hasta")
            loComandoSeleccionar.AppendLine("	")
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Recepciones'						AS	Operacion,	")
            loComandoSeleccionar.AppendLine(" 		Recepciones.Documento				AS	Documento,")
            loComandoSeleccionar.AppendLine(" 		Renglones_Recepciones.Cod_Art		AS	Cod_Art,")
            loComandoSeleccionar.AppendLine("		Operaciones_Lotes.Cod_Lot				AS	Lote, 	")
            loComandoSeleccionar.AppendLine(" 		Recepciones.Fec_Ini					AS	Fec_Ini, 	")
            loComandoSeleccionar.AppendLine(" 		Renglones_Recepciones.Renglon		AS	Renglon, 	 ")
            loComandoSeleccionar.AppendLine(" 		Renglones_Recepciones.Cod_Alm		AS	Cod_Alm, 	")
            'loComandoSeleccionar.AppendLine(" 		0.0									AS	CanRng_Sal, ")
            'loComandoSeleccionar.AppendLine(" 		Renglones_Recepciones.Can_Art1		AS	CanRng_Ent, ")
            loComandoSeleccionar.AppendLine("		0.0									AS	CanLte_Sal, ")
            loComandoSeleccionar.AppendLine(" 		Operaciones_Lotes.Cantidad			AS	CanLte_Ent,	")
            loComandoSeleccionar.AppendLine(" 		'Entrada'							AS	Tipo,")
            loComandoSeleccionar.AppendLine(" 		Articulos.Saldo				        AS	Saldo ")
            loComandoSeleccionar.AppendLine("FROM Recepciones")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Recepciones ON Renglones_Recepciones.Documento = Recepciones.Documento")
            loComandoSeleccionar.AppendLine("	JOIN @tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Recepciones.Cod_Art ")
            loComandoSeleccionar.AppendLine("	JOIN Operaciones_Lotes ON Operaciones_Lotes.Num_Doc = Recepciones.Documento")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Doc = 'Recepciones'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Tip_Ope = 'Entrada'")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Ren_Ori = Renglones_Recepciones.Renglon")
            loComandoSeleccionar.AppendLine("		AND Operaciones_Lotes.Cod_Art = Renglones_Recepciones.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE Recepciones.Status IN ('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("   AND Recepciones.Fec_Ini <= @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("	AND Renglones_Recepciones.Cod_Alm BETWEEN @lcCodAlm_Desde AND @lcCodAlm_Hasta")
            loComandoSeleccionar.AppendLine("   AND Operaciones_Lotes.Cod_Lot BETWEEN @lcCodLot_Desde AND @lcCodLot_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("---- Crea un índice para acelerar las siguientes operaciones")
            loComandoSeleccionar.AppendLine("CREATE CLUSTERED INDEX PK_Fecha ON #curTemporal(Fec_Ini)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("---- Calcula los saldos iniciales")
            loComandoSeleccionar.AppendLine("UPDATE #curTemporal")
            loComandoSeleccionar.AppendLine("SET Saldo = S.Saldo")
            loComandoSeleccionar.AppendLine("FROM (	SELECT	Cod_Art, Cod_Alm, Lote,")
            loComandoSeleccionar.AppendLine("			SUM(CanLte_Ent - CanLte_Sal) AS Saldo		")
            loComandoSeleccionar.AppendLine("		FROM	#curTemporal")
            loComandoSeleccionar.AppendLine("		WHERE	#curTemporal.Fec_Ini < @ldFecha_Desde")
            loComandoSeleccionar.AppendLine("		GROUP BY Cod_Art, Cod_Alm, Lote) AS S")
            loComandoSeleccionar.AppendLine("WHERE #curTemporal.Fec_Ini >= @ldFecha_Desde")
            loComandoSeleccionar.AppendLine("	AND	#curTemporal.Cod_Art = S.Cod_Art ")
            loComandoSeleccionar.AppendLine("	AND	#curTemporal.Lote = S.Lote ")
            loComandoSeleccionar.AppendLine("	AND	#curTemporal.Cod_Alm = S.Cod_Alm ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	#curTemporal.Saldo									AS Inicial,		")
            loComandoSeleccionar.AppendLine("		#curTemporal.Saldo									AS Saldo,		")
            loComandoSeleccionar.AppendLine("		#curTemporal.Operacion								AS Operacion,	")
            loComandoSeleccionar.AppendLine("		#curTemporal.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("		#curTemporal.Cod_Art								AS Cod_Art,")
            loComandoSeleccionar.AppendLine("		#curTemporal.Lote									AS Lote, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.Fec_Ini								AS Fec_Ini, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.Renglon								AS Renglon, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.Cod_Alm								AS Cod_Alm, 	")
            'loComandoSeleccionar.AppendLine("		#curTemporal.CanRng_Sal								AS CanRng_Sal, 	")
            'loComandoSeleccionar.AppendLine("		#curTemporal.CanRng_Ent								AS CanRng_Ent, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.CanLte_Sal								AS Can_Sal, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.CanLte_Ent								AS Can_Ent, 	")
            loComandoSeleccionar.AppendLine("		#curTemporal.Tipo									AS Tipo,")
            loComandoSeleccionar.AppendLine("		Articulos.Nom_Art									AS Nom_Art,		")
            loComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFecha_Desde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFecha_Hasta AS DATE),103))	AS Fecha,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Art_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodArt_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Art FROM Articulos WHERE Cod_Art = @lcCodArt_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Art_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodDep_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Dep_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodDep_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Dep FROM Departamentos WHERE Cod_Dep = @lcCodDep_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Dep_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSec_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Sec_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodSec_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Sec FROM Secciones WHERE Cod_Sec = @lcCodSec_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Sec_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Desde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Desde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Alm_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcCodAlm_Hasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Alm  FROM Almacenes  WHERE Cod_Alm = @lcCodAlm_Hasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				AS Alm_Hasta,")
            loComandoSeleccionar.AppendLine("		CONCAT(@lcCodLot_Desde, ' - ', (CASE WHEN @lcCodLot_Hasta = 'zzzzzzz' THEN '' ELSE @lcCodLot_Hasta END))	AS Lotes")
            loComandoSeleccionar.AppendLine("FROM #curTemporal")
            loComandoSeleccionar.AppendLine("   JOIN Articulos ")
            loComandoSeleccionar.AppendLine("       ON Articulos.Cod_Art = #curTemporal.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE #curTemporal.Fec_Ini >= @ldFecha_Desde")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Art ASC,  Fec_Ini ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #curTemporal")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Calcula el saldo de cada movimiento por artículo
            '-------------------------------------------------------------------------------------------------------
            Dim lcArticulo As String = ""
            Dim lnSaldo As Decimal = 0D
            For Each loRenglon As DataRow In laDatosReporte.Tables(0).Rows

                If Trim(loRenglon("Cod_Art")).ToLower() <> lcArticulo Then
                    lcArticulo = Trim(loRenglon("Cod_Art")).ToLower()
                    lnSaldo = CDec(loRenglon("Inicial"))
                End If

                lnSaldo = lnSaldo + CDec(loRenglon("Can_Ent")) - CDec(loRenglon("Can_Sal"))
                loRenglon("Saldo") = lnSaldo

            Next loRenglon


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rHMovimientos_Inventarios", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rHMovimientos_Inventarios.ReportSource = loObjetoReporte

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
' JJD: 30/09/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 12/05/09: Ordenamiento 
'-------------------------------------------------------------------------------------------'
' AAP: 29/06/09: Filtro "Sucursal"
'-------------------------------------------------------------------------------------------'
' JJD: 24/07/09: Se desgloso por origen. Ajuste, Traslados, Devoluciones Clientes, etc.
'-------------------------------------------------------------------------------------------'
' CMS: 19/08/09: Verificacion de registros y filtro cod_ubi
'-------------------------------------------------------------------------------------------'
' RJG: 14/03/10: Corecciones varias: no mostraba todos los documentos, corrección en		'
'				 Traslados, Compras, Entregas. Ajustados Status de documentos de Compra y	'
'				 Venta (agregados los Afectados y Procesados).								'
'-------------------------------------------------------------------------------------------'
' RJG: 08/12/10: Ajustado Estatus de Facturas de Venta: Ahora omite las facturas de venta	'
'				 pendientes e incluye las confirmadas.										'
'-------------------------------------------------------------------------------------------'
' RJG: 15/05/12: Se cambió el SP y el SELECT de movimientos para que aplique correctamente	'
'				 los filtros de almacen y sucursal (los cálculos no eran correctos).		'
'-------------------------------------------------------------------------------------------'
' RJG: 14/06/12: Se cambió SELECT de movimientos para que no use el SP.						'
'-------------------------------------------------------------------------------------------'
