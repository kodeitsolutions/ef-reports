'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rHMovimientos_Inventarios_Costos"
'-------------------------------------------------------------------------------------------'
Partial Class rHMovimientos_Inventarios_Costos
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
            Dim lcParametro10Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            '-------------------------------------------------------------------------------------------'
            ' Invocacion del Store Procedure
            '-------------------------------------------------------------------------------------------'
			Dim lcAlmacenTraslado AS String = goServicios.mObtenerCampoFormatoSQL(cusadministrativo.goAlmacen.pcAlmacenTransito)
			
            loComandoSeleccionar.AppendLine("DECLARE @tmpArticulos AS TABLE(Cod_Art CHAR(30), Saldo DECIMAL(28,10)) ;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT 	INTO @tmpArticulos ")
            loComandoSeleccionar.AppendLine("SELECT		Cod_Art, 0")
            loComandoSeleccionar.AppendLine("FROM		Articulos")
            loComandoSeleccionar.AppendLine("WHERE		Cod_Art 	BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Dep 	BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Sec 	BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Mar 	BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Cla 	BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Tip 	BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Pro 	BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Uni1	BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Ubi		BETWEEN " & lcParametro12Desde & " AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())


            '-------------------------------------------------------------------------------------------'
            ' El código del Almacén de Tránsito se requiere para los Traslados entre Almacénes          '
            ' Confirmados o Procesados.	                                                                '
            '-------------------------------------------------------------------------------------------'
            Dim lcAlmacenTransito As String
            lcAlmacenTransito = goServicios.mObtenerCampoFormatoSQL(goOpciones.mObtener("CODALMTRA", "C"))

            Select Case lcParametro10Desde

                Case "'Todos'"

                    ' Select de la tabla de Ajustes
                    loComandoSeleccionar.AppendLine("SELECT		'Ajustes'								AS	Operacion, 	")
                    loComandoSeleccionar.AppendLine("			Ajustes.Documento						AS	Documento, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Art				AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine("			Ajustes.Fec_Ini							AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Renglon				AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine("			'No Aplica'								AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Alm				AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine("			(CASE WHEN Renglones_Ajustes.Tipo = 'Salida' THEN Renglones_Ajustes.Can_Art1 ELSE 0.0 END)		AS	Can_Sal, ")
                    loComandoSeleccionar.AppendLine("			(CASE WHEN Renglones_Ajustes.Tipo = 'Entrada'  THEN Renglones_Ajustes.Can_Art1 ELSE 0.0 END)		AS	Can_Ent, ")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Tipo					AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Ajustes.Cos_Pro1              AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Ajustes.Cos_Ult1              AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Ajustes.Cod_Suc							AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
                    loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
                    loComandoSeleccionar.AppendLine("FROM		Ajustes")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Ajustes.Status					=	'Confirmado' ")
                    loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Tipo			IN	('Entrada', 'Salida') ")
                    loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Fec_Ini					<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Cod_Alm		BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Cod_Suc					BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro9Hasta)

                    ' Union con Select de la tabla de Traslados en las Salidas
                    loComandoSeleccionar.AppendLine("UNION ALL ")
                    loComandoSeleccionar.AppendLine("SELECT		'Traslados'								AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Documento						AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cod_Art				AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Fec_Ini						AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Renglon				AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			'No Aplica'								AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Alm_Ori						AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Can_Art1			AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Salida'								AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cos_Pro1            AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cos_Ult1            AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Cod_Suc						AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
                    loComandoSeleccionar.AppendLine("FROM		Traslados")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Traslados.Status		IN ('Confirmado', 'Procesado')	")
                    loComandoSeleccionar.AppendLine(" 		AND Traslados.Fec_Ini				<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Traslados.Alm_Ori		BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Traslados.Cod_Suc				BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)
					
                    ' Union con Select de la tabla de Traslados en las Entradas
                    loComandoSeleccionar.AppendLine("UNION ALL ")
                    loComandoSeleccionar.AppendLine("SELECT		'Traslados'								AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Documento						AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cod_Art				AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Fec_Ini						AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Renglon				AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			'No Aplica'								AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			CASE Traslados.Status					")
                    loComandoSeleccionar.AppendLine(" 			    WHEN 'Confirmado'	THEN " & lcAlmacenTransito)
                    loComandoSeleccionar.AppendLine(" 			    WHEN 'Procesado'	THEN Traslados.Alm_Des")
                    loComandoSeleccionar.AppendLine(" 			END										AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Can_Art1			AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Entrada'								AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cos_Pro1            AS	Cos_Pro,	") 
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cos_Ult1            AS	Cos_Ult,	") 
                    loComandoSeleccionar.AppendLine(" 			Traslados.Cod_Suc						AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
                    loComandoSeleccionar.AppendLine("FROM		Traslados")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Traslados.Status			IN ('Confirmado', 'Procesado')	")
                    loComandoSeleccionar.AppendLine(" 		AND Traslados.Fec_Ini			<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND (CASE Traslados.Status					")
                    loComandoSeleccionar.AppendLine(" 			    WHEN 'Confirmado'	THEN " & lcAlmacenTransito)
                    loComandoSeleccionar.AppendLine(" 			    WHEN 'Procesado'	THEN Traslados.Alm_Des")
                    loComandoSeleccionar.AppendLine(" 			END)						BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Traslados.Cod_Suc			BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

                    ' Union con Select de la tabla de Entregas
                    loComandoSeleccionar.AppendLine("UNION ALL ")
                    loComandoSeleccionar.AppendLine("SELECT		'Entregas'							AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Entregas.Documento					AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cod_Art			AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Entregas.Fec_Ini					AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Renglon			AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			Entregas.Cod_Cli					AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cod_Alm			AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Can_Art1			AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Salida'							AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cos_Pro1         AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cos_Ult1         AS	Cos_Ult,	")
					loComandoSeleccionar.AppendLine(" 			Entregas.Cod_Suc					AS	Cod_Suc,	")
					loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo		")
                    loComandoSeleccionar.AppendLine("FROM		Entregas")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Entregas ON Renglones_Entregas.Documento = Entregas.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Entregas.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Entregas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine(" 		AND Entregas.Fec_Ini				<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Renglones_Entregas.Cod_Alm		BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Entregas.Cod_Suc				BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

                    ' Union con Select de la tabla de Facturas
                    loComandoSeleccionar.AppendLine("UNION ALL ")
                    loComandoSeleccionar.AppendLine("SELECT		'Facturas'							AS	Operacion, ")
                    loComandoSeleccionar.AppendLine(" 			Facturas.Documento					AS	Documento, ")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cod_Art			AS	Cod_Art, ")
                    loComandoSeleccionar.AppendLine(" 			Facturas.Fec_Ini					AS	Fec_Ini, ")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Renglon			AS	Renglon, ")
                    loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Cli					AS	Cliente, ")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cod_Alm			AS	Cod_Alm, ")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Can_Art1			AS	Can_Sal, ")
                    loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Ent, ")
                    loComandoSeleccionar.AppendLine(" 			'Salida'							AS	Tipo, ")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cos_Pro1         AS	Cos_Pro, ") 
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cos_Ult1         AS	Cos_Ult, ") 
					loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Suc					AS	Cod_Suc, ")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo ")
                    loComandoSeleccionar.AppendLine("FROM		Facturas")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Facturas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine("		AND Renglones_Facturas.Tip_Ori		<>	'Entregas' ")
                    loComandoSeleccionar.AppendLine("		AND Facturas.Fec_Ini				<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine("		AND Renglones_Facturas.Cod_Alm		BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Suc				BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

                    ' Union con Select de la tabla de Recepciones
                    loComandoSeleccionar.AppendLine("UNION ALL ")
                    loComandoSeleccionar.AppendLine("SELECT		'Recepciones'						AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Recepciones.Documento				AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cod_Art		AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Recepciones.Fec_Ini					AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Renglon		AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			Recepciones.Cod_Pro					AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cod_Alm		AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Can_Art1		AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Entrada'							AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cos_Pro1      AS  Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cos_Ult1      AS  Cos_Ult,	")
					loComandoSeleccionar.AppendLine(" 			Recepciones.Cod_Suc					AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo ")
                    loComandoSeleccionar.AppendLine("FROM		Recepciones")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Recepciones ON Renglones_Recepciones.Documento = Recepciones.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Recepciones.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Recepciones.Status			IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine("		AND Recepciones.Fec_Ini				<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine("		AND Renglones_Recepciones.Cod_Alm	BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("		AND Recepciones.Cod_Suc				BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

                    ' Union con Select de la tabla de Compras
                    loComandoSeleccionar.AppendLine("UNION ALL ")
                    loComandoSeleccionar.AppendLine("SELECT		'Compras'							AS	Operacion,	")
                    loComandoSeleccionar.AppendLine("			Compras.Documento					AS	Documento,	")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Art			AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine("			Compras.Fec_Ini						AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Renglon			AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine("			Compras.Cod_Pro						AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Alm			AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine("			0.0									AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Can_Art1			AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine("			'Entrada'							AS	Tipo,		")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Cos_Pro1          AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Cos_Ult1          AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine("			Compras.Cod_Suc						AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo		")
                    loComandoSeleccionar.AppendLine("FROM		Compras")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Compras.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Compras.Status				IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine(" 		AND Renglones_Compras.Tip_Ori	<>	'Recepciones' ")
                    loComandoSeleccionar.AppendLine("		AND Compras.Fec_Ini				<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine("		AND Renglones_Compras.Cod_Alm	BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("		AND Compras.Cod_Suc				BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

                    ' Union con Select de la tabla de Devoluciones_Clientes
                    loComandoSeleccionar.AppendLine(" UNION ALL ")
                    loComandoSeleccionar.AppendLine(" SELECT	'Dev_Cli'								AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Documento			AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cod_Art				AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Fec_Ini			AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Renglon				AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Cod_Cli			AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cod_Alm				AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Can_Art1			AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Entrada'								AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cos_Pro1            AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cos_Ult1            AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Cod_Suc			AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
                    loComandoSeleccionar.AppendLine("FROM		Devoluciones_Clientes")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_DClientes ON Renglones_DClientes.Documento = Devoluciones_Clientes.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_DClientes.Cod_Art ")
                    loComandoSeleccionar.AppendLine(" WHERE			Devoluciones_Clientes.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine(" 			AND Devoluciones_Clientes.Fec_Ini		<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 			AND Renglones_DClientes.Cod_Alm			BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 			AND Devoluciones_Clientes.Cod_Suc		BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 				AND " & lcParametro9Hasta)

                    ' Union con Select de la tabla de Devoluciones_Proveedores
                    loComandoSeleccionar.AppendLine("UNION ALL ")
                    loComandoSeleccionar.AppendLine("SELECT		'Dev_Pro'								AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Documento		AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cod_Art			AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Fec_Ini		AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Renglon			AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Cod_Pro		AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cod_Alm			AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Can_Art1			AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Salida'								AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cos_Pro1         AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cos_Ult1         AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Cod_Suc        AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
                    loComandoSeleccionar.AppendLine("FROM		Devoluciones_Proveedores")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_DProveedores ON Renglones_DProveedores.Documento = Devoluciones_Proveedores.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_DProveedores.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Devoluciones_Proveedores.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Proveedores.Fec_Ini	<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Renglones_DProveedores.Cod_Alm		BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Proveedores.Cod_Suc					BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)
                    
                Case "'Ajustes_Entrada'"

                    ' Select de la tabla de Ajustes solo para las Entradas
                    loComandoSeleccionar.AppendLine("SELECT		'Ajustes'								AS	Operacion,	")
                    loComandoSeleccionar.AppendLine("			Ajustes.Documento						AS	Documento,	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Art				AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine("			Ajustes.Fec_Ini							AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Renglon				AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine("			'No Aplica'								AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Alm				AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine("			0.0										AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Can_Art1				AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Tipo					AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Ajustes.Cos_Pro1              AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Ajustes.Cos_Ult1              AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Ajustes.Cod_Suc							AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
                    loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
                    loComandoSeleccionar.AppendLine("FROM		Ajustes")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Ajustes.Status					=	'Confirmado' ")
                    loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Tipo			=	'Entrada' ")
                    loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Fec_Ini					<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Cod_Alm		BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Cod_Suc					BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro9Hasta)

                Case "'Ajustes_Salida'"

                    ' Select de la tabla de Ajustes solo para las Salidas
                    loComandoSeleccionar.AppendLine("SELECT		'Ajustes'								AS	Operacion,	")
                    loComandoSeleccionar.AppendLine("			Ajustes.Documento						AS	Documento,	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Art				AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine("			Ajustes.Fec_Ini							AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Renglon				AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine("			'No Aplica'								AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Alm				AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Can_Art1				AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine("			0.0										AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Tipo					AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Ajustes.Cos_Pro1              AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Ajustes.Cos_Ult1              AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Ajustes.Cod_Suc							AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
                    loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
                    loComandoSeleccionar.AppendLine("FROM		Ajustes")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Ajustes.Status					=	'Confirmado' ")
                    loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Tipo			=	'Salida' ")
                    loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Fec_Ini					<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Cod_Alm		BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Cod_Suc					BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro9Hasta)

                Case "'Traslados_Salida'"

                    ' Union con Select de la tabla de Traslados en las Salidas
                    loComandoSeleccionar.AppendLine("SELECT		'Traslados'								AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Documento						AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cod_Art				AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Fec_Ini						AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Renglon				AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			'No Aplica'								AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Alm_Ori						AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Can_Art1			AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Salida'								AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cos_Pro1            AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cos_Ult1            AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Cod_Suc						AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
                    loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
                    loComandoSeleccionar.AppendLine("FROM		Traslados")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Traslados.Status		IN ('Confirmado', 'Procesado')	")
                    loComandoSeleccionar.AppendLine(" 		AND Traslados.Fec_Ini				<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Traslados.Alm_Ori		BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Traslados.Cod_Suc				BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

                Case "'Traslados_Entrada'"

                    ' Union con Select de la tabla de Traslados en las Entradas
                    loComandoSeleccionar.AppendLine("SELECT		'Traslados'								AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Documento						AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cod_Art				AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Fec_Ini						AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Renglon				AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			'No Aplica'								AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			CASE Traslados.Status					")
                    loComandoSeleccionar.AppendLine(" 			    WHEN 'Confirmado'	THEN " & lcAlmacenTransito)
                    loComandoSeleccionar.AppendLine(" 			    WHEN 'Procesado'	THEN Traslados.Alm_Des")
                    loComandoSeleccionar.AppendLine(" 			END										AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Can_Art1			AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Entrada'								AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cos_Pro1            AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Traslados.Cos_Ult1            AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Traslados.Cod_Suc						AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
                    loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
                    loComandoSeleccionar.AppendLine("FROM		Traslados")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Traslados ON Renglones_Traslados.Documento = Traslados.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Traslados.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Traslados.Status			IN ('Confirmado', 'Procesado')	")
                    loComandoSeleccionar.AppendLine(" 		AND Traslados.Fec_Ini			<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND (CASE Traslados.Status					")
                    loComandoSeleccionar.AppendLine(" 			    WHEN 'Confirmado'	THEN " & lcAlmacenTransito)
                    loComandoSeleccionar.AppendLine(" 			    WHEN 'Procesado'	THEN Traslados.Alm_Des")
                    loComandoSeleccionar.AppendLine(" 			END)						BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Traslados.Cod_Suc			BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

                Case "'Facturas_Venta'"

                    ' Union con Select de la tabla de Facturas
                    loComandoSeleccionar.AppendLine("SELECT		'Facturas'							AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Facturas.Documento					AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cod_Art			AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Facturas.Fec_Ini					AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Renglon			AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Cli					AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cod_Alm			AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Can_Art1			AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Salida'							AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cos_Pro1         AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cos_Ult1         AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Suc					AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo		")
                    loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
                    loComandoSeleccionar.AppendLine("FROM		Facturas")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Facturas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine("		AND Renglones_Facturas.Tip_Ori		<>	'Entregas' ")
                    loComandoSeleccionar.AppendLine("		AND Facturas.Fec_Ini				<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine("		AND Renglones_Facturas.Cod_Alm		BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Suc				BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

                Case "'Facturas_Compra'"

                    ' Union con Select de la tabla de Compras
                    loComandoSeleccionar.AppendLine("SELECT		'Compras'							AS	Operacion,	")
                    loComandoSeleccionar.AppendLine("			Compras.Documento					AS	Documento,	")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Art			AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine("			Compras.Fec_Ini						AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Renglon			AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine("			Compras.Cod_Pro						AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Alm			AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine("			0.0									AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Can_Art1			AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine("			'Entrada'							AS	Tipo,		")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Cos_Pro1          AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine("			Renglones_Compras.Cos_Ult1          AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Compras.Cod_Suc						AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo	")
                    loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
                    loComandoSeleccionar.AppendLine("FROM		Compras")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Compras.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Compras.Status				IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine(" 		AND Renglones_Compras.Tip_Ori	<>	'Recepciones' ")
                    loComandoSeleccionar.AppendLine("		AND Compras.Fec_Ini				<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine("		AND Renglones_Compras.Cod_Alm	BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("		AND Compras.Cod_Suc				BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

                Case "'Notas_Entrega'"

                    ' Union con Select de la tabla de Entregas
                    loComandoSeleccionar.AppendLine("SELECT		'Entregas'							AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Entregas.Documento					AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cod_Art			AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Entregas.Fec_Ini					AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Renglon			AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			Entregas.Cod_Cli					AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cod_Alm			AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Can_Art1			AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Salida'							AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cos_Pro1         AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cos_Ult1         AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Entregas.Cod_Suc					AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo		")
                    loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
                    loComandoSeleccionar.AppendLine("FROM		Entregas")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Entregas ON Renglones_Entregas.Documento = Entregas.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Entregas.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Entregas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine(" 		AND Entregas.Fec_Ini				<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Renglones_Entregas.Cod_Alm		BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Entregas.Cod_Suc				BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

                Case "'Notas_Recepcion'"

                    ' Union con Select de la tabla de Recepciones
                    loComandoSeleccionar.AppendLine("SELECT		'Recepciones'						AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Recepciones.Documento				AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cod_Art		AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Recepciones.Fec_Ini					AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Renglon		AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			Recepciones.Cod_Pro					AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cod_Alm		AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Can_Art1		AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Entrada'							AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cos_Pro1      AS  Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cos_Ult1      AS  Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Recepciones.Cod_Suc					AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo	    ")
                    loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
                    loComandoSeleccionar.AppendLine("FROM		Recepciones")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_Recepciones ON Renglones_Recepciones.Documento = Recepciones.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Recepciones.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Recepciones.Status			IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine("		AND Recepciones.Fec_Ini				<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine("		AND Renglones_Recepciones.Cod_Alm	BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine("		AND Recepciones.Cod_Suc				BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)

                Case "'Devolucion_Ventas'"

                    ' Union con Select de la tabla de Devoluciones_Clientes
                    loComandoSeleccionar.AppendLine(" SELECT	'Dev_Cli'								AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Documento			AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cod_Art				AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Fec_Ini			AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Renglon				AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Cod_Cli			AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cod_Alm				AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Can_Art1			AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Entrada'								AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cos_Pro1            AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cos_Ult1            AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Cod_Suc			AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
                    loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
                    loComandoSeleccionar.AppendLine("FROM		Devoluciones_Clientes")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_DClientes ON Renglones_DClientes.Documento = Devoluciones_Clientes.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_DClientes.Cod_Art ")
                    loComandoSeleccionar.AppendLine(" WHERE			Devoluciones_Clientes.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine(" 			AND Devoluciones_Clientes.Fec_Ini		<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 			AND Renglones_DClientes.Cod_Alm			BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 			AND Devoluciones_Clientes.Cod_Suc		BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 				AND " & lcParametro9Hasta)

                Case "'Devolucion_Compras'"

                    ' Union con Select de la tabla de Devoluciones_Proveedores
                    loComandoSeleccionar.AppendLine("SELECT		'Dev_Pro'								AS	Operacion,	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Documento		AS	Documento,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cod_Art			AS	Cod_Art, 	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Fec_Ini		AS	Fec_Ini, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Renglon			AS	Renglon, 	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Cod_Pro		AS	Cliente, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cod_Alm			AS	Cod_Alm, 	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Can_Art1			AS	Can_Sal, 	")
                    loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Ent, 	")
                    loComandoSeleccionar.AppendLine(" 			'Salida'								AS	Tipo,		")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cos_Pro1         AS	Cos_Pro,	")
                    loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cos_Ult1         AS	Cos_Ult,	")
                    loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Cod_Suc		AS	Cod_Suc,	")
                    loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo		")
                    loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
                    loComandoSeleccionar.AppendLine("FROM		Devoluciones_Proveedores")
                    loComandoSeleccionar.AppendLine("	JOIN	Renglones_DProveedores ON Renglones_DProveedores.Documento = Devoluciones_Proveedores.Documento")
                    loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_DProveedores.Cod_Art ")
                    loComandoSeleccionar.AppendLine("WHERE		Devoluciones_Proveedores.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
                    loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Proveedores.Fec_Ini	<=	" & lcParametro1Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Renglones_DProveedores.Cod_Alm		BETWEEN " & lcParametro2Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
                    loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Proveedores.Cod_Suc	BETWEEN " & lcParametro9Desde)
                    loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)

            End Select

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- Crea un índice para acelerar las siguientes operaciones")
            loComandoSeleccionar.AppendLine("CREATE CLUSTERED INDEX PK_Fecha ON #curTemporal(Fec_Ini)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- Calcula los saldos iniciales")
            loComandoSeleccionar.AppendLine("UPDATE		#curTemporal")
            loComandoSeleccionar.AppendLine("SET			Saldo = S.Saldo")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT	Cod_Art, Cod_Alm, Cod_Suc,")
            loComandoSeleccionar.AppendLine("					SUM(Can_Ent - Can_Sal) As Saldo		")
            loComandoSeleccionar.AppendLine("			FROM	#curTemporal")
            loComandoSeleccionar.AppendLine("			WHERE	#curTemporal.Fec_Ini < " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			GROUP BY Cod_Art, Cod_Alm, Cod_Suc")
            loComandoSeleccionar.AppendLine("		)	AS S")
            loComandoSeleccionar.AppendLine("WHERE 		#curTemporal.Fec_Ini >= " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("	AND		#curTemporal.Cod_Art = S.Cod_Art  ")
            loComandoSeleccionar.AppendLine("	AND		#curTemporal.Cod_Alm = S.Cod_Alm ")
            loComandoSeleccionar.AppendLine("	AND		#curTemporal.Cod_Suc = S.Cod_Suc")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		#curTemporal.Saldo									AS Inicial,		")
            loComandoSeleccionar.AppendLine("			#curTemporal.Saldo									AS Saldo,		")
            loComandoSeleccionar.AppendLine("			#curTemporal.Operacion								AS Operacion,	")
            loComandoSeleccionar.AppendLine("			#curTemporal.Documento								AS Documento,	")
            loComandoSeleccionar.AppendLine("			#curTemporal.Cod_Art								AS Cod_Art, 	")
            loComandoSeleccionar.AppendLine("			#curTemporal.Fec_Ini								AS Fec_Ini, 	")
            loComandoSeleccionar.AppendLine("			#curTemporal.Renglon								AS Renglon, 	")
            loComandoSeleccionar.AppendLine("			#curTemporal.Cliente								AS Cliente, 	")
            loComandoSeleccionar.AppendLine("			#curTemporal.Cod_Alm								AS Cod_Alm, 	")
            loComandoSeleccionar.AppendLine("			#curTemporal.Can_Sal								AS Can_Sal, 	")
            loComandoSeleccionar.AppendLine("			#curTemporal.Can_Ent								AS Can_Ent, 	")
            loComandoSeleccionar.AppendLine("			#curTemporal.Tipo									AS Tipo,		")
            loComandoSeleccionar.AppendLine("			#curTemporal.Cos_Pro								AS Cos_Pro,		")
            loComandoSeleccionar.AppendLine("			#curTemporal.Cos_Ult								AS Cos_Ult,		")
            'loComandoSeleccionar.AppendLine("			(#curTemporal.Cos_Pro * #curTemporal.Can_Sal)		AS Cos_Pro_Sal, ")
            'loComandoSeleccionar.AppendLine("			(#curTemporal.Cos_Pro * #curTemporal.Can_Ent)		AS Cos_Pro_Ent, ")
            'loComandoSeleccionar.AppendLine("			(	(#curTemporal.Cos_Pro * #curTemporal.Can_Ent)					")
            'loComandoSeleccionar.AppendLine("			  - (#curTemporal.Cos_Pro * #curTemporal.Can_Sal))	AS Cos_Pro_Tot,")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art									AS Nom_Art		")
            loComandoSeleccionar.AppendLine("FROM		#curTemporal")
            loComandoSeleccionar.AppendLine("	JOIN	Articulos ")
            loComandoSeleccionar.AppendLine("		ON	Articulos.Cod_Art = #curTemporal.Cod_Art ")
            loComandoSeleccionar.AppendLine("WHERE 		#curTemporal.Fec_Ini >= " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #curTemporal")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

        '-------------------------------------------------------------------------------------------------------
        ' Calcula el saldo de cada movimiento por artículo
        '-------------------------------------------------------------------------------------------------------
			Dim lcArticulo	As String = ""
			Dim lnSaldo		As Decimal= 0D
			For Each loRenglon As DataRow In laDatosReporte.Tables(0).Rows
				
				If Trim(loRenglon("Cod_Art")).ToLower() <> lcArticulo Then 
					lcArticulo	= Trim(loRenglon("Cod_Art")).ToLower()
					lnSaldo		= CDec(loRenglon("Inicial"))	
				End If
				
				lnSaldo		= lnSaldo + CDec(loRenglon("Can_Ent")) - CDec(loRenglon("Can_Sal"))
				loRenglon("Saldo") = lnSaldo
				
			Next loRenglon
			
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rHMovimientos_Inventarios_Costos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrHMovimientos_Inventarios_Costos.ReportSource = loObjetoReporte

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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 12/12/14: Codigo inicial, a partir de rHMovimientos_Inventarios.                     '
'-------------------------------------------------------------------------------------------'
