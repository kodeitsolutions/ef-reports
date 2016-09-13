'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Inventarios"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Inventarios
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
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine("DECLARE @tmpArticulos AS TABLE(Cod_Art CHAR(30), Saldo DECIMAL(28,10)) ;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT 	INTO @tmpArticulos ")
            loComandoSeleccionar.AppendLine("SELECT		Cod_Art, 0")
            loComandoSeleccionar.AppendLine("FROM		Articulos")
            loComandoSeleccionar.AppendLine("WHERE		Cod_Art 	BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Dep 	BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Sec 	BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Mar 	BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Cla 	BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Tip 	BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Pro 	BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Uni1	BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("		AND	Cod_Ubi		BETWEEN " & lcParametro10Desde & " AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())


            ' Select de la tabla de Ajustes
            loComandoSeleccionar.AppendLine("SELECT		'Ajustes'								AS	Operacion, 		 ")
            loComandoSeleccionar.AppendLine("			Ajustes.Documento						AS	Documento, 		 ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Art				AS	Cod_Art, 		 ")
            loComandoSeleccionar.AppendLine("			Ajustes.Fec_Ini							AS	Fec_Ini, 		 ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Renglon				AS	Renglon, 		 ")
            loComandoSeleccionar.AppendLine("                'No Aplica'								AS	Cliente, 		 ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Cod_Alm				AS	Cod_Alm, 		 ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Renglones_Ajustes.Tipo = 'Salida' THEN Renglones_Ajustes.Can_Art1 ELSE 0.0 END)		AS	Can_Sal, 	 ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Renglones_Ajustes.Tipo = 'Entrada'  THEN Renglones_Ajustes.Can_Art1 ELSE 0.0 END)		AS	Can_Ent, 	 ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes.Tipo					AS	Tipo, ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Ajustes.Cos_Pro1              AS	Cos_Pro,	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Ajustes.Cos_Ult1              AS	Cos_Ult, ")
            loComandoSeleccionar.AppendLine(" 			Ajustes.Cod_Suc							AS	Cod_Suc, ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo, ")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))				AS Costo_Stock_Inicial, ")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))				AS Unitario_Stock_Inicial, ")
            loComandoSeleccionar.AppendLine("			Renglones_Ajustes.cod_uni				AS unidad ")
            loComandoSeleccionar.AppendLine("INTO		#curTemporal ")
            loComandoSeleccionar.AppendLine("FROM		Ajustes")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Ajustes ON Renglones_Ajustes.Documento = Ajustes.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Ajustes.Cod_Art ")
            loComandoSeleccionar.AppendLine("WHERE		Ajustes.Status					=	'Confirmado' ")
            loComandoSeleccionar.AppendLine(" 		AND	Renglones_Ajustes.Tipo			IN	('Entrada', 'Salida') ")
            loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Fec_Ini					<=	" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 		AND	Ajustes.Cod_Suc					BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			AND	" & lcParametro8Hasta)


            ' Union con Select de la tabla de Entregas
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT		'Entregas'							AS	Operacion,		")
            loComandoSeleccionar.AppendLine(" 			Entregas.Documento					AS	Documento,		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cod_Art			AS	Cod_Art, 		")
            loComandoSeleccionar.AppendLine(" 			Entregas.Fec_Ini					AS	Fec_Ini, 		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Renglon			AS	Renglon, 		")
            loComandoSeleccionar.AppendLine(" 			Entregas.Cod_Cli					AS	Cliente, 		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cod_Alm			AS	Cod_Alm, 		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Can_Art1			AS	Can_Sal, 		")
            loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Ent, 		")
            loComandoSeleccionar.AppendLine("                'Salida'							AS	Tipo,			")
            loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cos_Pro1         AS	Cos_Pro,		")
            loComandoSeleccionar.AppendLine(" 			Renglones_Entregas.Cos_Ult1         AS	Cos_Ult,		")
            loComandoSeleccionar.AppendLine(" 			Entregas.Cod_Suc					AS	Cod_Suc,		")
            loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo,	")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))			AS Costo_Stock_Inicial,	")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))			AS Unitario_Stock_Inicial,	")
            loComandoSeleccionar.AppendLine("			Renglones_Entregas.cod_uni			AS unidad	")
            loComandoSeleccionar.AppendLine("FROM		Entregas")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Entregas ON Renglones_Entregas.Documento = Entregas.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Entregas.Cod_Art ")
            loComandoSeleccionar.AppendLine("WHERE		Entregas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 		AND Entregas.Fec_Ini				<=	" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Entregas.Cod_Suc				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro8Hasta)

            ' Union con Select de la tabla de Facturas
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT		'Facturas'							AS	Operacion, 	")
            loComandoSeleccionar.AppendLine(" 			Facturas.Documento					AS	Documento, 	")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cod_Art			AS	Cod_Art, 	")
            loComandoSeleccionar.AppendLine(" 			Facturas.Fec_Ini					AS	Fec_Ini, 	")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Renglon			AS	Renglon, 	")
            loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Cli					AS	Cliente, 	")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cod_Alm			AS	Cod_Alm, 	")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Can_Art1			AS	Can_Sal, 	")
            loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Ent, 	")
            loComandoSeleccionar.AppendLine("                'Salida'							AS	Tipo, 	")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cos_Pro1         AS	Cos_Pro, 	")
            loComandoSeleccionar.AppendLine(" 			Renglones_Facturas.Cos_Ult1         AS	Cos_Ult, 	")
            loComandoSeleccionar.AppendLine(" 			Facturas.Cod_Suc					AS	Cod_Suc, 	")
            loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo,	")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))			AS Costo_Stock_Inicial,	")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))			AS Unitario_Stock_Inicial,	")
            loComandoSeleccionar.AppendLine("			Renglones_Facturas.cod_uni			AS unidad	")
            loComandoSeleccionar.AppendLine("FROM		Facturas")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art ")
            loComandoSeleccionar.AppendLine("WHERE		Facturas.Status					IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Facturas.Tip_Ori		<>	'Entregas' ")
            loComandoSeleccionar.AppendLine("		AND Facturas.Fec_Ini				<=	" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Suc				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro8Hasta)

            ' Union con Select de la tabla de Recepciones
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT		'Recepciones'						AS	Operacion,	 ")
            loComandoSeleccionar.AppendLine(" 			Recepciones.Documento				AS	Documento,	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cod_Art		AS	Cod_Art, 	 ")
            loComandoSeleccionar.AppendLine(" 			Recepciones.Fec_Ini					AS	Fec_Ini, 	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Renglon		AS	Renglon, 	 ")
            loComandoSeleccionar.AppendLine(" 			Recepciones.Cod_Pro					AS	Cliente, 	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cod_Alm		AS	Cod_Alm, 	 ")
            loComandoSeleccionar.AppendLine(" 			0.0									AS	Can_Sal, 	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Can_Art1		AS	Can_Ent, 	 ")
            loComandoSeleccionar.AppendLine("                'Entrada'						AS	Tipo,		 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cos_Pro1      AS  Cos_Pro,	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_Recepciones.Cos_Ult1      AS  Cos_Ult,	 ")
            loComandoSeleccionar.AppendLine(" 			Recepciones.Cod_Suc					AS	Cod_Suc,	 ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo, ")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))			AS Costo_Stock_Inicial, ")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))			AS Unitario_Stock_Inicial, ")
            loComandoSeleccionar.AppendLine("			Renglones_Recepciones.cod_uni		AS unidad	  ")
            loComandoSeleccionar.AppendLine("FROM		Recepciones")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Recepciones ON Renglones_Recepciones.Documento = Recepciones.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Recepciones.Cod_Art ")
            loComandoSeleccionar.AppendLine("WHERE		Recepciones.Status			IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine("		AND Recepciones.Fec_Ini				<=	" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Recepciones.Cod_Suc				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro8Hasta)

            ' Union con Select de la tabla de Compras
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT		'Compras'							AS	Operacion,	 ")
            loComandoSeleccionar.AppendLine("			Compras.Documento					AS	Documento,	 ")
            loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Art			AS	Cod_Art, 	 ")
            loComandoSeleccionar.AppendLine("			Compras.Fec_Ini						AS	Fec_Ini, 	 ")
            loComandoSeleccionar.AppendLine("			Renglones_Compras.Renglon			AS	Renglon, 	 ")
            loComandoSeleccionar.AppendLine("			Compras.Cod_Pro						AS	Cliente, 	 ")
            loComandoSeleccionar.AppendLine("			Renglones_Compras.Cod_Alm			AS	Cod_Alm, 	 ")
            loComandoSeleccionar.AppendLine("			0.0									AS	Can_Sal, 	 ")
            loComandoSeleccionar.AppendLine("			Renglones_Compras.Can_Art1			AS	Can_Ent, 	 ")
            loComandoSeleccionar.AppendLine("                'Entrada'						AS	Tipo,		 ")
            loComandoSeleccionar.AppendLine("			Renglones_Compras.Cos_Pro1          AS	Cos_Pro,	 ")
            loComandoSeleccionar.AppendLine("			Renglones_Compras.Cos_Ult1          AS	Cos_Ult,	 ")
            loComandoSeleccionar.AppendLine("			Compras.Cod_Suc						AS	Cod_Suc,	 ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				        AS	Saldo, ")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))			AS  Costo_Stock_Inicial, ")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))			AS  Unitario_Stock_Inicial, ")
            loComandoSeleccionar.AppendLine("			Renglones_Compras.cod_uni			AS  unidad	 ")
            loComandoSeleccionar.AppendLine("FROM		Compras")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Compras ON Renglones_Compras.Documento = Compras.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_Compras.Cod_Art ")
            loComandoSeleccionar.AppendLine("WHERE		Compras.Status				IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 		AND Renglones_Compras.Tip_Ori	<>	'Recepciones' ")
            loComandoSeleccionar.AppendLine("		AND Compras.Fec_Ini				<=	" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Compras.Cod_Suc				BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro8Hasta)

            ' Union con Select de la tabla de Devoluciones_Clientes
            loComandoSeleccionar.AppendLine(" UNION ALL ")
            loComandoSeleccionar.AppendLine(" SELECT		'Dev_Cli'							AS	Operacion,	 ")
            loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Documento			AS	Documento,	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cod_Art				AS	Cod_Art, 	 ")
            loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Fec_Ini			AS	Fec_Ini, 	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Renglon				AS	Renglon, 	 ")
            loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Cod_Cli			AS	Cliente, 	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cod_Alm				AS	Cod_Alm, 	 ")
            loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Sal, 	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Can_Art1			AS	Can_Ent, 	 ")
            loComandoSeleccionar.AppendLine("                'Entrada'							AS	Tipo,		 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cos_Pro1            AS	Cos_Pro,	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DClientes.Cos_Ult1            AS	Cos_Ult,	 ")
            loComandoSeleccionar.AppendLine(" 			Devoluciones_Clientes.Cod_Suc			AS	Cod_Suc,	 ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo, ")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))				AS Costo_Stock_Inicial, ")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))				AS Unitario_Stock_Inicial, ")
            loComandoSeleccionar.AppendLine("			Renglones_DClientes.cod_uni				AS unidad ")
            loComandoSeleccionar.AppendLine("FROM		Devoluciones_Clientes")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_DClientes ON Renglones_DClientes.Documento = Devoluciones_Clientes.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_DClientes.Cod_Art ")
            loComandoSeleccionar.AppendLine(" WHERE			Devoluciones_Clientes.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 			AND Devoluciones_Clientes.Fec_Ini		<=	" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Devoluciones_Clientes.Cod_Suc		BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro8Hasta)

            ' Union con Select de la tabla de Devoluciones_Proveedores
            loComandoSeleccionar.AppendLine("UNION ALL ")
            loComandoSeleccionar.AppendLine("SELECT		'Dev_Pro'								AS	Operacion,	 ")
            loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Documento		AS	Documento,	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cod_Art			AS	Cod_Art, 	 ")
            loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Fec_Ini		AS	Fec_Ini, 	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Renglon			AS	Renglon, 	 ")
            loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Cod_Pro		AS	Cliente, 	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cod_Alm			AS	Cod_Alm, 	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Can_Art1			AS	Can_Sal, 	 ")
            loComandoSeleccionar.AppendLine(" 			0.0										AS	Can_Ent, 	 ")
            loComandoSeleccionar.AppendLine("                'Salida'							AS	Tipo,		 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cos_Pro1         AS	Cos_Pro,	 ")
            loComandoSeleccionar.AppendLine(" 			Renglones_DProveedores.Cos_Ult1         AS	Cos_Ult,	 ")
            loComandoSeleccionar.AppendLine(" 			Devoluciones_Proveedores.Cod_Suc        AS	Cod_Suc,	 ")
            loComandoSeleccionar.AppendLine(" 			Articulos.Saldo				            AS	Saldo, ")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))				AS Costo_Stock_Inicial, ")
            loComandoSeleccionar.AppendLine("			CAST(0 AS DECIMAL(28,10))				AS Unitario_Stock_Inicial, ")
            loComandoSeleccionar.AppendLine("			Renglones_DProveedores.cod_uni			AS unidad	 ")
            loComandoSeleccionar.AppendLine("FROM		Devoluciones_Proveedores")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_DProveedores ON Renglones_DProveedores.Documento = Devoluciones_Proveedores.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	@tmpArticulos AS Articulos ON Articulos.Cod_Art = Renglones_DProveedores.Cod_Art ")
            loComandoSeleccionar.AppendLine("WHERE		Devoluciones_Proveedores.Status		IN	('Confirmado', 'Afectado', 'Procesado') ")
            loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Proveedores.Fec_Ini	<=	" & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Devoluciones_Proveedores.Cod_Suc					BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro8Hasta)


            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- Crea un índice para acelerar las siguientes operaciones")
            loComandoSeleccionar.AppendLine("CREATE CLUSTERED INDEX PK_Fecha ON #curTemporal(Fec_Ini)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- Calcula los saldos iniciales")
            loComandoSeleccionar.AppendLine("UPDATE		#curTemporal")
            loComandoSeleccionar.AppendLine("SET			Saldo = S.Saldo,")
            loComandoSeleccionar.AppendLine("   			Costo_Stock_Inicial = s.costo")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT	Cod_Art, Cod_Alm, Cod_Suc,")
            loComandoSeleccionar.AppendLine("					SUM(Can_Ent - Can_Sal) As Saldo,		")
            loComandoSeleccionar.AppendLine("					SUM(ROUND(Can_Ent*cos_ult - Can_Sal*cos_ult,2)) AS costo")
            loComandoSeleccionar.AppendLine("			FROM	#curTemporal")
            loComandoSeleccionar.AppendLine("			WHERE	#curTemporal.Fec_Ini < " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			GROUP BY Cod_Art, Cod_Alm, Cod_Suc")
            loComandoSeleccionar.AppendLine("		)	AS S")
            loComandoSeleccionar.AppendLine("WHERE 		#curTemporal.Fec_Ini >= " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("	AND		#curTemporal.Cod_Art = S.Cod_Art  ")
            loComandoSeleccionar.AppendLine("	AND		#curTemporal.Cod_Suc = S.Cod_Suc")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT		#curTemporal.Cod_Art									AS Cod_Art,  ")
            loComandoSeleccionar.AppendLine("			Articulos.Nom_Art										AS Nom_Art, ")
            loComandoSeleccionar.AppendLine("			#curTemporal.unidad										AS unidad, ")
            loComandoSeleccionar.AppendLine("			SUM(#curTemporal.Saldo)									AS Stock_Inicial, ")
            loComandoSeleccionar.AppendLine("			SUM(#curTemporal.Costo_stock_inicial)					AS Costo_stock_inicial, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion  ")
            loComandoSeleccionar.AppendLine("					WHEN 'Compras' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("					WHEN 'Recepciones' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Pro' THEN -#curTemporal.Can_Sal ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)												AS Compras, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion  ")
            loComandoSeleccionar.AppendLine("					WHEN 'Compras' THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Recepciones' THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Pro' THEN -#curTemporal.Can_Sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)/ ")
            loComandoSeleccionar.AppendLine("			CASE  ")
            loComandoSeleccionar.AppendLine("				SUM(CASE #curTemporal.operacion  ")
            loComandoSeleccionar.AppendLine("						WHEN 'Compras' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("						WHEN 'Recepciones' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("						WHEN 'Dev_Pro' THEN -#curTemporal.Can_Sal ")
            loComandoSeleccionar.AppendLine("						ELSE 0 ")
            loComandoSeleccionar.AppendLine("					END) ")
            loComandoSeleccionar.AppendLine("				WHEN 0 THEN 1 ")
            loComandoSeleccionar.AppendLine("				ELSE ")
            loComandoSeleccionar.AppendLine("					SUM(CASE #curTemporal.operacion  ")
            loComandoSeleccionar.AppendLine("							WHEN 'Compras' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("							WHEN 'Recepciones' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("							WHEN 'Dev_Pro' THEN -#curTemporal.Can_Sal ")
            loComandoSeleccionar.AppendLine("							ELSE 0 ")
            loComandoSeleccionar.AppendLine("						END) ")
            loComandoSeleccionar.AppendLine("			END													AS Unitario_Compras, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion  ")
            loComandoSeleccionar.AppendLine("					WHEN 'Compras' THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Recepciones' THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Pro' THEN -#curTemporal.Can_Sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)												AS Total_Compras, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Entrada' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)												AS Entradas, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Entrada' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)/ ")
            loComandoSeleccionar.AppendLine("			CASE  ")
            loComandoSeleccionar.AppendLine("				SUM(CASE  ")
            loComandoSeleccionar.AppendLine("						WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("							AND #curTemporal.Tipo ='Entrada' ")
            loComandoSeleccionar.AppendLine("						THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("						ELSE 0 ")
            loComandoSeleccionar.AppendLine("					END) ")
            loComandoSeleccionar.AppendLine("				WHEN 0 THEN 1 ")
            loComandoSeleccionar.AppendLine("				ELSE ")
            loComandoSeleccionar.AppendLine("					SUM(CASE  ")
            loComandoSeleccionar.AppendLine("							WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("								AND #curTemporal.Tipo ='Entrada' ")
            loComandoSeleccionar.AppendLine("							THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("							ELSE 0 ")
            loComandoSeleccionar.AppendLine("						END) ")
            loComandoSeleccionar.AppendLine("			END													AS Unitario_Entradas, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Entrada' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)												AS Total_Entradas, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion ")
            loComandoSeleccionar.AppendLine("					WHEN 'Facturas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("					WHEN 'Entregas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Cli'	THEN -#curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)												AS Ventas, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion ")
            loComandoSeleccionar.AppendLine("					WHEN 'Facturas' THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Entregas' THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Cli'	THEN -#curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)/ ")
            loComandoSeleccionar.AppendLine("			CASE ")
            loComandoSeleccionar.AppendLine("				SUM(CASE #curTemporal.operacion ")
            loComandoSeleccionar.AppendLine("						WHEN 'Facturas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("						WHEN 'Entregas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("						WHEN 'Dev_Cli'	THEN -#curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("						ELSE 0 ")
            loComandoSeleccionar.AppendLine("					END) ")
            loComandoSeleccionar.AppendLine("				WHEN 0 THEN 1 ")
            loComandoSeleccionar.AppendLine("				ELSE ")
            loComandoSeleccionar.AppendLine("					SUM(CASE #curTemporal.operacion ")
            loComandoSeleccionar.AppendLine("							WHEN 'Facturas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("							WHEN 'Entregas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("							WHEN 'Dev_Cli'	THEN -#curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("							ELSE 0 ")
            loComandoSeleccionar.AppendLine("						END) ")
            loComandoSeleccionar.AppendLine("			END														AS Unitario_Ventas, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion ")
            loComandoSeleccionar.AppendLine("					WHEN 'Facturas' THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Entregas' THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Cli'	THEN -#curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)												AS total_Ventas, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Salida' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)												AS Salidas, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Salida' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)/ ")
            loComandoSeleccionar.AppendLine("			CASE ")
            loComandoSeleccionar.AppendLine("				SUM(CASE  ")
            loComandoSeleccionar.AppendLine("						WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("							AND #curTemporal.Tipo ='Salida' ")
            loComandoSeleccionar.AppendLine("						THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("						ELSE 0 ")
            loComandoSeleccionar.AppendLine("					END) ")
            loComandoSeleccionar.AppendLine("				WHEN 0 THEN 1 ")
            loComandoSeleccionar.AppendLine("				ELSE ")
            loComandoSeleccionar.AppendLine("					SUM(CASE  ")
            loComandoSeleccionar.AppendLine("							WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("								AND #curTemporal.Tipo ='Salida' ")
            loComandoSeleccionar.AppendLine("							THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("							ELSE 0 ")
            loComandoSeleccionar.AppendLine("						END) ")
            loComandoSeleccionar.AppendLine("			END													AS Unitario_Salidas, ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Salida' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)													AS Total_Salidas, ")
            loComandoSeleccionar.AppendLine("			SUM(#curTemporal.Costo_stock_inicial)+ ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion  ")
            loComandoSeleccionar.AppendLine("					WHEN 'Compras' THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Recepciones' THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Pro' THEN -#curTemporal.Can_Sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)+ ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Entrada' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)- ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion ")
            loComandoSeleccionar.AppendLine("					WHEN 'Facturas' THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Entregas' THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Cli'	THEN -#curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)- ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Salida' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)												AS Costo_Stock_Final, ")
            loComandoSeleccionar.AppendLine("			(SUM(#curTemporal.Costo_stock_inicial)+ ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion  ")
            loComandoSeleccionar.AppendLine("					WHEN 'Compras' THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Recepciones' THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Pro' THEN -#curTemporal.Can_Sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)+ ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Entrada' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)- ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion ")
            loComandoSeleccionar.AppendLine("					WHEN 'Facturas' THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Entregas' THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Cli'	THEN -#curTemporal.Can_Ent*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)- ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Salida' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_sal*#curTemporal.Cos_Ult ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END))/ ")
            loComandoSeleccionar.AppendLine("			CASE ")
            loComandoSeleccionar.AppendLine("				SUM(#curTemporal.Saldo)+ ")
            loComandoSeleccionar.AppendLine("				SUM(CASE #curTemporal.operacion  ")
            loComandoSeleccionar.AppendLine("						WHEN 'Compras' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("						WHEN 'Recepciones' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("						WHEN 'Dev_Pro' THEN -#curTemporal.Can_Sal ")
            loComandoSeleccionar.AppendLine("						ELSE 0 ")
            loComandoSeleccionar.AppendLine("					END) + ")
            loComandoSeleccionar.AppendLine("				SUM(CASE  ")
            loComandoSeleccionar.AppendLine("						WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("							AND #curTemporal.Tipo ='Entrada' ")
            loComandoSeleccionar.AppendLine("						THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("						ELSE 0 ")
            loComandoSeleccionar.AppendLine("					END) - ")
            loComandoSeleccionar.AppendLine("				SUM(CASE #curTemporal.operacion ")
            loComandoSeleccionar.AppendLine("						WHEN 'Facturas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("						WHEN 'Entregas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("						WHEN 'Dev_Cli'	THEN -#curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("						ELSE 0 ")
            loComandoSeleccionar.AppendLine("					END) - ")
            loComandoSeleccionar.AppendLine("				SUM(CASE  ")
            loComandoSeleccionar.AppendLine("						WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("							AND #curTemporal.Tipo ='Salida' ")
            loComandoSeleccionar.AppendLine("						THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("						ELSE 0 ")
            loComandoSeleccionar.AppendLine("					END) ")
            loComandoSeleccionar.AppendLine("				WHEN 0 THEN 1 ")
            loComandoSeleccionar.AppendLine("				ELSE ")
            loComandoSeleccionar.AppendLine("					(SUM(#curTemporal.Saldo)+ ")
            loComandoSeleccionar.AppendLine("					SUM(CASE #curTemporal.operacion  ")
            loComandoSeleccionar.AppendLine("							WHEN 'Compras' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("							WHEN 'Recepciones' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("							WHEN 'Dev_Pro' THEN -#curTemporal.Can_Sal ")
            loComandoSeleccionar.AppendLine("							ELSE 0 ")
            loComandoSeleccionar.AppendLine("						END) + ")
            loComandoSeleccionar.AppendLine("					SUM(CASE  ")
            loComandoSeleccionar.AppendLine("							WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("								AND #curTemporal.Tipo ='Entrada' ")
            loComandoSeleccionar.AppendLine("							THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("							ELSE 0 ")
            loComandoSeleccionar.AppendLine("						END) - ")
            loComandoSeleccionar.AppendLine("					SUM(CASE #curTemporal.operacion ")
            loComandoSeleccionar.AppendLine("							WHEN 'Facturas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("							WHEN 'Entregas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("							WHEN 'Dev_Cli'	THEN -#curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("							ELSE 0 ")
            loComandoSeleccionar.AppendLine("						END) - ")
            loComandoSeleccionar.AppendLine("					SUM(CASE  ")
            loComandoSeleccionar.AppendLine("							WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("								AND #curTemporal.Tipo ='Salida' ")
            loComandoSeleccionar.AppendLine("							THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("							ELSE 0 ")
            loComandoSeleccionar.AppendLine("						END)) ")
            loComandoSeleccionar.AppendLine("			END															AS Unitario_final, ")
            loComandoSeleccionar.AppendLine("			SUM(#curTemporal.Saldo)+ ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion  ")
            loComandoSeleccionar.AppendLine("					WHEN 'Compras' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("					WHEN 'Recepciones' THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Pro' THEN -#curTemporal.Can_Sal ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END) + ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Entrada' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END) - ")
            loComandoSeleccionar.AppendLine("			SUM(CASE #curTemporal.operacion ")
            loComandoSeleccionar.AppendLine("					WHEN 'Facturas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("					WHEN 'Entregas' THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("					WHEN 'Dev_Cli'	THEN -#curTemporal.Can_Ent ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END) - ")
            loComandoSeleccionar.AppendLine("			SUM(CASE  ")
            loComandoSeleccionar.AppendLine("					WHEN #curTemporal.operacion = 'Ajustes'  ")
            loComandoSeleccionar.AppendLine("						AND #curTemporal.Tipo ='Salida' ")
            loComandoSeleccionar.AppendLine("					THEN #curTemporal.Can_sal ")
            loComandoSeleccionar.AppendLine("					ELSE 0 ")
            loComandoSeleccionar.AppendLine("				END)													AS Inventario_final, ")
            loComandoSeleccionar.AppendLine("            'Libro de Inventario al '+ ")
            loComandoSeleccionar.AppendLine("			RIGHT('00'+CAST(DAY(CAST(" & lcParametro1Hasta & "	AS DATETIME)) AS VARCHAR),2)+'/'+ ")
            loComandoSeleccionar.AppendLine("			RIGHT('00'+CAST(MONTH(CAST(" & lcParametro1Hasta & "	AS DATETIME)) AS VARCHAR),2)+'/'+ ")
            loComandoSeleccionar.AppendLine("			CAST(YEAR(CAST(" & lcParametro1Hasta & "	AS DATETIME)) AS VARCHAR) AS titulo ")
            loComandoSeleccionar.AppendLine("FROM		#curTemporal")
            loComandoSeleccionar.AppendLine("	JOIN	Articulos ")
            loComandoSeleccionar.AppendLine("		ON	Articulos.Cod_Art = #curTemporal.Cod_Art ")
            loComandoSeleccionar.AppendLine("WHERE 		#curTemporal.Fec_Ini >= " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("GROUP BY	#curTemporal.Cod_Art, Articulos.nom_art,#curTemporal.unidad")
            loComandoSeleccionar.AppendLine("ORDER BY	#curTemporal.Cod_Art ASC")
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Inventarios", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLibro_Inventarios.ReportSource = loObjetoReporte
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


            'Selección de opcion por excel (Microsoft Excel - xls)
            If (Me.Request.QueryString("salida").ToLower() = "xls") Then
                ' Ruta donde se creara temporalmente el archivo
                Dim lcFileName As String = Server.MapPath("~\Contabilidad\Temporales\rLibro_Inventarios_" & Guid.NewGuid().ToString("N") & ".xls")
                ' Se exporta para crear el archivo temporal
                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, lcFileName)

                ' Se modifica el contenido del archivo
                Dim lcParametrosReporte As String = cusAplicacion.goReportes.mObtenerParametros(cusAplicacion.goReportes.paNombresParametros, cusAplicacion.goReportes.paParametrosIniciales, cusAplicacion.goReportes.paParametrosFinales)
                Me.mGenerarArchivoExcel(lcFileName, laDatosReporte.Tables(0), lcParametrosReporte)

                ' Se coloca en la respuesta para decargar
                Me.Response.Clear()
                Me.Response.Buffer = True
                Me.Response.AppendHeader("content-disposition", "attachment; filename=rLibro_Inventarios.xls")
                Me.Response.ContentType = "application/excel"
                Me.Response.WriteFile(lcFileName, True)
                Me.Response.Write(Space(30))
                Me.Response.Flush()
                Me.Response.Close()

                Me.Response.End()

            Else

                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "Este reporte fue diseñado solo para mostrar la vista ""Microsoft Excel - xls"". ", _
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
    Private Sub mGenerarArchivoExcel(ByVal lcNombreArchivo As String, ByVal loDatos As DataTable, ByVal lcParametrosReporte As String)

        Dim lnDecimalesMonto As Integer = goOpciones.pnDecimalesParaMonto
        Dim lnDecimalesCosto As Integer = goOpciones.pnDecimalesParaCosto
        Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
        Dim lnDecimalesPorcentaje As Integer = goOpciones.pnDecimalesParaPorcentaje

        Dim lcFormatoMontos As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesMonto)
        Dim lcFormatoCosto As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesCosto)

        Dim lcFormatoCantidad As String
        If (lnDecimalesCantidad > 0) Then
            lcFormatoCantidad = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesCantidad)
        Else
            lcFormatoCantidad = "###,###,###,###,##0"
        End If

        Dim lcFormatoPorcentaje As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesPorcentaje)

        '******************************************************************'
        ' Declaración de objetos de excel: IMPORTANTE liberar recursos al	'
        ' final usando el GARBAGE COLLECTOR y ReleaseComObject.			'
        '******************************************************************'
        Dim loExcel As Excel.Application = Nothing
        Dim laLibros As Excel.Workbooks = Nothing
        Dim loLibro As Excel.Workbook = Nothing
        Dim loHoja As Excel.Worksheet = Nothing
        Dim loCeldas As Excel.Range = Nothing
        Dim loRango As Excel.Range = Nothing

        Dim loFilas As Excel.Range = Nothing
        Dim loColumnas As Excel.Range = Nothing
        Dim loFormas As Excel.Shapes = Nothing
        Dim loImagen As Excel.Shape = Nothing
        Dim loFuente As Excel.Font = Nothing


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

            ' Formato por defecto de todas las celdas			
            loCeldas = loHoja.Range("A1:IV65536")
            'loCeldas = loHoja.Cells
            loCeldas.Clear()
            loFuente = loCeldas.Font
            loFuente.Size = 9
            loFuente.Name = "Tahoma"


            '******************************************************************'
            ' Encabezado de la hoja											'
            '******************************************************************'

            loRango = loHoja.Range("A1")
            loRango.Value = cusAplicacion.goEmpresa.pcNombre

            loRango = loHoja.Range("A2")
            loRango.Value = cusAplicacion.goEmpresa.pcRifEmpresa
            Dim loRenglon1 As DataRow = loDatos.Rows(0)
            loRango = loHoja.Range("B5:T5")
            loRango.Select()
            loRango.MergeCells = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.NumberFormat = "@"
            loRango.Value = CStr(loRenglon1("titulo")).Trim()
            loFuente = loRango.Font
            loFuente.Size = 14
            loFuente.Bold = True

            'Fecha y hora de creacion
            Dim ldFecha As DateTime = Date.Now()
            loRango = loHoja.Range("T1")
            loRango.NumberFormat = "@"
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            loRango.Value = ldFecha.ToString("dd/MM/yyyy")

            loRango = loHoja.Range("T2")
            loRango.NumberFormat = "@" 'La celda almacena un string
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            loRango.Value = ldFecha.ToString("hh:mm:ss tt")

            '' Parametros del reporte
            'loRango = loHoja.Range("B7:J7")
            'loRango.Select()
            'loRango.MergeCells = True
            'loRango.Value = lcParametrosReporte
            'loRango.WrapText = True
            'loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify


            Dim lnFilaActual As Integer = 9

            ''******************************************************************'
            '' Datos del Reporte												'
            ''******************************************************************'

            loRango = loHoja.Range("B" & lnFilaActual)
            loRango.Value = "Código"
            loRango = loHoja.Range("B" & (lnFilaActual) & ":B" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("C" & lnFilaActual)
            loRango.Value = "Nombre"
            loRango = loHoja.Range("C" & (lnFilaActual) & ":C" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("D" & lnFilaActual)
            loRango.Value = "Uni"
            loRango = loHoja.Range("D" & (lnFilaActual) & ":D" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("E" & lnFilaActual)
            loRango.Value = "Stock" & vbLf & "Inicial"
            loRango = loHoja.Range("E" & (lnFilaActual) & ":E" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("F" & lnFilaActual)
            loRango.Value = "Compras"
            loRango = loHoja.Range("F" & (lnFilaActual) & ":F" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("G" & lnFilaActual)
            loRango.Value = "Costo" & vbLf & "Unitario"
            loRango = loHoja.Range("G" & (lnFilaActual) & ":G" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("H" & lnFilaActual)
            loRango.Value = "Total en" & vbLf & "Compras"
            loRango = loHoja.Range("H" & (lnFilaActual) & ":H" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("I" & lnFilaActual)
            loRango.Value = "Entradas"
            loRango = loHoja.Range("I" & (lnFilaActual) & ":I" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("J" & lnFilaActual)
            loRango.Value = "Costo" & vbLf & "Unitario"
            loRango = loHoja.Range("J" & (lnFilaActual) & ":J" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("K" & lnFilaActual)
            loRango.Value = "Total en" & vbLf & "Entradas"
            loRango = loHoja.Range("K" & (lnFilaActual) & ":K" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("L" & lnFilaActual)
            loRango.Value = "Ventas"
            loRango = loHoja.Range("L" & (lnFilaActual) & ":L" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("M" & lnFilaActual)
            loRango.Value = "Costo" & vbLf & "Unitario"
            loRango = loHoja.Range("M" & (lnFilaActual) & ":M" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("N" & lnFilaActual)
            loRango.Value = "Total en" & vbLf & "Ventas"
            loRango = loHoja.Range("N" & (lnFilaActual) & ":N" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("O" & lnFilaActual)
            loRango.Value = "Salidas"
            loRango = loHoja.Range("O" & (lnFilaActual) & ":O" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("P" & lnFilaActual)
            loRango.Value = "Costo" & vbLf & "Unitario"
            loRango = loHoja.Range("P" & (lnFilaActual) & ":P" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("Q" & lnFilaActual)
            loRango.Value = "Total en" & vbLf & "Salidas"
            loRango = loHoja.Range("Q" & (lnFilaActual) & ":Q" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("R" & lnFilaActual)
            loRango.Value = "Costo Stock" & vbLf & "Final"
            loRango = loHoja.Range("R" & (lnFilaActual) & ":R" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("S" & lnFilaActual)
            loRango.Value = "Costo" & vbLf & "Unitario"
            loRango = loHoja.Range("S" & (lnFilaActual) & ":S" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            loRango = loHoja.Range("T" & lnFilaActual)
            loRango.Value = "Inventario" & vbLf & "Final"
            loRango = loHoja.Range("T" & (lnFilaActual) & ":T" & (lnFilaActual))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)



            loRango = loHoja.Range("B" & lnFilaActual & ":T" & lnFilaActual)
            loFuente = loRango.Font
            loFuente.Bold = True
            'loFuente.Color = RGB(255, 255, 255)
            loRango.Interior.Color = RGB(179, 179, 179)

            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            Dim lnFilaInicio As Integer = lnFilaActual
            For lnRenglon As Integer = 0 To loDatos.Rows.Count - 1
                Dim loRenglon As DataRow = loDatos.Rows(lnRenglon)

                lnFilaActual += 1

                'Código
                loRango = loHoja.Range("B" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Cod_Art")).Trim()
                loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                loRango = loHoja.Range("B" & (lnFilaActual) & ":B" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

                'Nombre de Articulo
                loRango = loHoja.Range("C" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Nom_Art")).Trim()
                loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                loRango = loHoja.Range("C" & (lnFilaActual) & ":C" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Unidad
                loRango = loHoja.Range("D" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Unidad")).Trim()
                loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                loRango = loHoja.Range("D" & (lnFilaActual) & ":D" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Stock Inicial
                loRango = loHoja.Range("E" & lnFilaActual)
                loRango.NumberFormat = lcFormatoCantidad
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Stock_Inicial")), lnDecimalesCantidad)
                loRango = loHoja.Range("E" & (lnFilaActual) & ":E" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Compras
                loRango = loHoja.Range("F" & lnFilaActual)
                loRango.NumberFormat = lcFormatoCantidad
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Compras")), lnDecimalesCantidad)
                loRango = loHoja.Range("F" & (lnFilaActual) & ":F" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Unitario 
                loRango = loHoja.Range("G" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Unitario_Compras")), lnDecimalesCantidad)
                loRango = loHoja.Range("G" & (lnFilaActual) & ":G" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Total Compras
                loRango = loHoja.Range("H" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Total_Compras")), lnDecimalesCantidad)
                loRango = loHoja.Range("H" & (lnFilaActual) & ":H" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Entradas
                loRango = loHoja.Range("I" & lnFilaActual)
                loRango.NumberFormat = lcFormatoCantidad
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Entradas")), lnDecimalesCantidad)
                loRango = loHoja.Range("I" & (lnFilaActual) & ":I" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Unitario 
                loRango = loHoja.Range("J" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Unitario_Entradas")), lnDecimalesCantidad)
                loRango = loHoja.Range("J" & (lnFilaActual) & ":J" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Total Entradas
                loRango = loHoja.Range("K" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Total_Entradas")), lnDecimalesCantidad)
                loRango = loHoja.Range("K" & (lnFilaActual) & ":K" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Ventas
                loRango = loHoja.Range("L" & lnFilaActual)
                loRango.NumberFormat = lcFormatoCantidad
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Ventas")), lnDecimalesCantidad)
                loRango = loHoja.Range("L" & (lnFilaActual) & ":L" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Unitario 
                loRango = loHoja.Range("M" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Unitario_Ventas")), lnDecimalesCantidad)
                loRango = loHoja.Range("M" & (lnFilaActual) & ":M" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Total Ventas
                loRango = loHoja.Range("N" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("total_Ventas")), lnDecimalesCantidad)
                loRango = loHoja.Range("N" & (lnFilaActual) & ":N" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Salidas
                loRango = loHoja.Range("O" & lnFilaActual)
                loRango.NumberFormat = lcFormatoCantidad
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Salidas")), lnDecimalesCantidad)
                loRango = loHoja.Range("O" & (lnFilaActual) & ":O" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                ' Unitario
                loRango = loHoja.Range("P" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Unitario_Salidas")), lnDecimalesCantidad)
                loRango = loHoja.Range("P" & (lnFilaActual) & ":P" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Total Salidas
                loRango = loHoja.Range("Q" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Total_Salidas")), lnDecimalesCantidad)
                loRango = loHoja.Range("Q" & (lnFilaActual) & ":Q" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Costo de Inventario final
                loRango = loHoja.Range("R" & lnFilaActual)
                loRango.NumberFormat = lcFormatoCantidad
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Costo_Stock_Final")), lnDecimalesCantidad)
                loRango = loHoja.Range("R" & (lnFilaActual) & ":R" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                ' Unitario
                loRango = loHoja.Range("S" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Unitario_final")), lnDecimalesCantidad)
                loRango = loHoja.Range("S" & (lnFilaActual) & ":S" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


                'Inventario Final
                loRango = loHoja.Range("T" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Inventario_final")), lnDecimalesCantidad)
                loRango = loHoja.Range("T" & (lnFilaActual) & ":T" & (lnFilaActual))
                loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)



            Next lnRenglon

            Dim lnTotal As Integer = loDatos.Rows.Count

            loRango = loHoja.Range("B" & (lnFilaInicio + 1) & ":T" & (lnFilaInicio + lnTotal))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            Dim lnDesde As Integer = lnFilaInicio
            Dim lnHasta As Integer = lnFilaInicio + lnTotal

            lnFilaInicio += lnTotal + 2
            loRango = loHoja.Range("C" & (lnFilaInicio))
            'loRango.MergeCells = True
            loRango.NumberFormat = "@"
            loRango.Value = "Total Articulos: " & lnTotal.ToString()
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft



            loRango = loHoja.Range("E" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUM(E" & lnDesde & ":E" & lnHasta & ")"

            loRango = loHoja.Range("F" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoCantidad
            loRango.Formula = "=SUM(F" & lnDesde & ":F" & lnHasta & ")"

            loRango = loHoja.Range("H" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUM(H" & lnDesde & ":H" & lnHasta & ")"

            loRango = loHoja.Range("I" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoCantidad
            loRango.Formula = "=SUM(I" & lnDesde & ":I" & lnHasta & ")"

            loRango = loHoja.Range("K" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUM(K" & lnDesde & ":K" & lnHasta & ")"

            loRango = loHoja.Range("L" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoCantidad
            loRango.Formula = "=SUM(L" & lnDesde & ":L" & lnHasta & ")"

            loRango = loHoja.Range("N" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUM(N" & lnDesde & ":N" & lnHasta & ")"

            loRango = loHoja.Range("O" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoCantidad
            loRango.Formula = "=SUM(O" & lnDesde & ":O" & lnHasta & ")"

            loRango = loHoja.Range("Q" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUM(Q" & lnDesde & ":Q" & lnHasta & ")"

            loRango = loHoja.Range("R" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUM(R" & lnDesde & ":R" & lnHasta & ")"

            loRango = loHoja.Range("T" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoCantidad
            loRango.Formula = "=SUM(T" & lnDesde & ":T" & lnHasta & ")"

            'loRango = loHoja.Range("F" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoPorcentaje
            'loRango.Formula = "=IF(D" & (lnFilaInicio) & ">0, E" & (lnFilaInicio) & "*100/D" & (lnFilaInicio) & ", 100)"

            'loRango = loHoja.Range("G" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            'loRango.Formula = "=SUM(G" & lnDesde & ":G" & lnHasta & ")"

            'loRango = loHoja.Range("H" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            'loRango.Formula = "=SUM(H" & lnDesde & ":H" & lnHasta & ")"

            'loRango = loHoja.Range("I" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoPorcentaje
            'loRango.Formula = "=IF(G" & (lnFilaInicio) & ">0, H" & (lnFilaInicio) & "*100/G" & (lnFilaInicio) & ", 100)"

            ''loRango = loHoja.Range("J" & (lnFilaInicio))
            ''loRango.NumberFormat = lcFormatoMontos
            ''loRango.Formula = "=IF(G" & (lnFilaInicio) & ">0, I" & (lnFilaInicio) & "*100/G" & (lnFilaInicio) & ", 100)"

            loRango = loHoja.Range("B" & (lnFilaInicio) & ":T" & (lnFilaInicio))
            loFuente = loRango.Font
            loFuente.Bold = True

            loFilas = loCeldas.Rows
            loFilas.AutoFit()

            loColumnas = loCeldas.Rows
            loColumnas.AutoFit()

            loRango = loHoja.Range("B1:B" & lnFilaInicio)
            loRango.ColumnWidth = 11

            loRango = loHoja.Range("C1:C" & lnFilaInicio)
            loRango.ColumnWidth = 45

            loRango = loHoja.Range("D1:D" & lnFilaInicio)
            loRango.ColumnWidth = 4

            loRango = loHoja.Range("E1:E" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("F1:F" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("G1:G" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("H1:H" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("I1:I" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("J1:J" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("K1:K" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("L1:L" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("M1:M" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("N1:N" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("O1:O" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("P1:P" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("Q1:Q" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("R1:R" & lnFilaInicio)
            loRango.ColumnWidth = 15

            loRango = loHoja.Range("S1:S" & lnFilaInicio)
            loRango.ColumnWidth = 12

            loRango = loHoja.Range("T1:T" & lnFilaInicio)
            loRango.ColumnWidth = 12


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


End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' EAG: 23/09/15: Codigo inicial                                                             '
'-------------------------------------------------------------------------------------------'
