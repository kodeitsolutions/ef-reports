'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rSeguimientos_lPedidos"
'-------------------------------------------------------------------------------------------'
Partial Class rSeguimientos_lPedidos 
    Inherits vis2Formularios.frmReporte
    
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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

			Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Prepara la tabla para los resultados finales")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpFinal(	-- Encabezado")
            loComandoSeleccionar.AppendLine("						Documento CHAR(10), Fec_Ini DATETIME, Fec_Fin DATETIME, Cod_Cli VARCHAR(10), ")
            loComandoSeleccionar.AppendLine("						Nom_Cli VARCHAR(100), Status CHAR(10), Mon_Net DECIMAL(28, 10), ")
            loComandoSeleccionar.AppendLine("						-- Sipping (también desde encabezado)")
            loComandoSeleccionar.AppendLine("						For_Env VARCHAR(MAX), Tip_Env VARCHAR(MAX), Dir_Ent VARCHAR(MAX), Grupo CHAR(2),")
            loComandoSeleccionar.AppendLine("						-- Grupo de Tiempos Logísticos")
            loComandoSeleccionar.AppendLine("						Tl_Renglon INT, Tl_Actividad VARCHAR(MAX), Tl_Inicio DATETIME, ")
            loComandoSeleccionar.AppendLine("						Tl_Hor_Ini VARCHAR(10), Tl_Comentario VARCHAR(MAX), Tl_Por_Eje DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("						-- Grupo de Tiempos Logísticos")
            loComandoSeleccionar.AppendLine("						Em_Renglon INT, Em_Tipo VARCHAR(MAX), Em_Cantidad DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("						Em_Peso DECIMAL(28,10), Em_Unidad VARCHAR(MAX), Em_Comentario VARCHAR(MAX),")
            loComandoSeleccionar.AppendLine("						-- Grupo de Tiempos Logísticos")
            loComandoSeleccionar.AppendLine("						Ta_Renglon INT, Ta_Transporte VARCHAR(MAX), Ta_Referencia VARCHAR(MAX), ")
            loComandoSeleccionar.AppendLine("						Ta_Inicio DATETIME, Ta_Hor_Ini VARCHAR(10), Ta_Cantidad DECIMAL(28,10), ")
            loComandoSeleccionar.AppendLine("						Ta_For_Env VARCHAR(MAX), Ta_Operador VARCHAR(MAX),")
            loComandoSeleccionar.AppendLine("						-- Grupo de Seguimientos Logísticos")
            loComandoSeleccionar.AppendLine("						Sl_Renglon INT, Sl_Fecha DATETIME, Sl_Hora CHAR(10), Sl_Contacto VARCHAR(MAX), Sl_Accion VARCHAR(MAX),")
            loComandoSeleccionar.AppendLine("						Sl_Medio VARCHAR(MAX), Sl_Comentario VARCHAR(MAX), Sl_Prioridad VARCHAR(MAX),")
            loComandoSeleccionar.AppendLine("						-- Grupo de Documentación")
            loComandoSeleccionar.AppendLine("						Do_Renglon INT, Do_Nom_Doc CHAR(100), Do_Archivo CHAR(300), Do_Ruta VARCHAR(MAX)")
            loComandoSeleccionar.AppendLine("						)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Índices para acelerar los JOINs")
            loComandoSeleccionar.AppendLine("CREATE NONCLUSTERED INDEX IX_tmpFinal_Documento_Grupo ON #tmpFinal(Documento, Grupo)")
            loComandoSeleccionar.AppendLine("CREATE NONCLUSTERED INDEX IX_tmpFinal_Documento ON #tmpFinal(Documento) ")
            loComandoSeleccionar.AppendLine("CREATE NONCLUSTERED INDEX IX_tmpFinal_Cod_Cli ON #tmpFinal(Cod_Cli) ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("-- Documentos de Origen.									*")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("SELECT		Pedidos.Documento				AS Documento,	")
            loComandoSeleccionar.AppendLine("			Pedidos.Fec_Ini					AS Fec_Ini,		")
            loComandoSeleccionar.AppendLine("			Pedidos.Fec_Fin					AS Fec_Fin,		")
            loComandoSeleccionar.AppendLine("			Pedidos.Cod_Cli					AS Cod_Cli,		")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Pedidos.Nom_Cli>'')					")
            loComandoSeleccionar.AppendLine("				THEN	Pedidos.Nom_Cli					")
            loComandoSeleccionar.AppendLine("				ELSE	Clientes.Nom_Cli					")
            loComandoSeleccionar.AppendLine("			END 							AS Nom_Cli,		")
            loComandoSeleccionar.AppendLine("			Pedidos.Status					AS Status,		")
            loComandoSeleccionar.AppendLine("			Pedidos.Mon_Net					AS Mon_Net, 	")
            loComandoSeleccionar.AppendLine("			Pedidos.For_Env					AS For_Env, 	")
            loComandoSeleccionar.AppendLine("			Pedidos.Tip_Env					AS Tip_Env, 	")
            loComandoSeleccionar.AppendLine("			Pedidos.Dir_Ent					AS Dir_Ent, 	")
            loComandoSeleccionar.AppendLine("			CAST(Pedidos.Tie_Log AS XML)	AS Tie_Log, 	")
            loComandoSeleccionar.AppendLine("			CAST(Pedidos.Embalaje AS XML)	AS Embalaje,	")
            loComandoSeleccionar.AppendLine("			CAST(Pedidos.Tra_Adi AS XML)	AS Tra_Adi, 	")
            loComandoSeleccionar.AppendLine("			CAST(Pedidos.Seg_Log AS XML)	AS Seg_Log  	")
            loComandoSeleccionar.AppendLine("INTO		#tmpPedidos")
            loComandoSeleccionar.AppendLine("FROM		Pedidos")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes")
            loComandoSeleccionar.AppendLine("		ON	Clientes.Cod_Cli = Pedidos.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE		Pedidos.Documento	BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Pedidos.Fec_Ini	BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Pedidos.Cod_Cli	BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND Pedidos.Cod_Ven	BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND Pedidos.Status		IN (" & lcParametro4Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Pedidos.Cod_Tra	BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("		AND Pedidos.Cod_Mon	BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("		AND Pedidos.Cod_For	BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("		AND Pedidos.Cod_Rev	BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("		AND Pedidos.Cod_Suc	BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("-- Tiempos Logísticos")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpFinal(	Documento, Fec_Ini, Fec_Fin, Cod_Cli, Nom_Cli, Status,")
            loComandoSeleccionar.AppendLine("						Mon_Net, For_Env, Tip_Env, Dir_Ent, Grupo,")
            loComandoSeleccionar.AppendLine("						Tl_Renglon, Tl_Actividad, Tl_Inicio, Tl_Hor_Ini, ")
            loComandoSeleccionar.AppendLine("						Tl_Comentario, Tl_Por_Eje)")
            loComandoSeleccionar.AppendLine("SELECT	#tmpPedidos.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Fec_Ini									AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Fec_Fin									AS Fec_Fin,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Cod_Cli									AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Nom_Cli									AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Status									AS Status,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Mon_Net									AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.For_Env									AS For_Env, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Tip_Env									AS Tip_Env, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Dir_Ent									AS Dir_Ent,")
            loComandoSeleccionar.AppendLine("		CAST('TL' AS CHAR(2))								AS Grupo,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@renglon)[1]',		'INT')				AS Tl_Renglon,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@actividad)[1]',	'VARCHAR(MAX)')		AS Tl_Actividad,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@inicio)[1]',		'DATETIME')			AS Tl_Inicio,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@hor_ini)[1]',		'VARCHAR(MAX)')		AS Tl_Hor_Ini,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@comentario)[1]',	'VARCHAR(MAX)')		AS Tl_Comentario,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@por_eje)[1]',		'DECIMAL(28,10)')	AS Tl_Por_Eje")
            loComandoSeleccionar.AppendLine("FROM	#tmpPedidos")
            loComandoSeleccionar.AppendLine("	CROSS APPLY Tie_Log.nodes('//elementos/elemento') AS T(C)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("-- Embalaje")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpFinal(	Documento, Fec_Ini, Fec_Fin, Cod_Cli, Nom_Cli, Status,")
            loComandoSeleccionar.AppendLine("						Mon_Net, For_Env, Tip_Env, Dir_Ent, Grupo,")
            loComandoSeleccionar.AppendLine("						Em_Renglon, Em_Tipo, Em_Cantidad, ")
            loComandoSeleccionar.AppendLine("						Em_Peso, Em_Unidad, Em_Comentario)")
            loComandoSeleccionar.AppendLine("SELECT	#tmpPedidos.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Fec_Ini									AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Fec_Fin									AS Fec_Fin,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Cod_Cli									AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Nom_Cli									AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Status									AS Status,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Mon_Net									AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.For_Env									AS For_Env, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Tip_Env									AS Tip_Env, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Dir_Ent									AS Dir_Ent,")
            loComandoSeleccionar.AppendLine("		CAST('EM' AS CHAR(2))								AS Grupo,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@renglon)[1]',		'INT')				AS Em_Renglon,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@tipo)[1]',			'VARCHAR(MAX)')		AS Em_Tipo,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@cantidad)[1]', 	'DECIMAL(28,10)')	AS Em_Cantidad,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@peso)[1]',			'DECIMAL(28,10)')	AS Em_Peso,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@unidad)[1]',		'VARCHAR(MAX)')		AS Em_Unidad,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@comentario)[1]',	'VARCHAR(MAX)')		AS Em_Comentario")
            loComandoSeleccionar.AppendLine("FROM	#tmpPedidos")
            loComandoSeleccionar.AppendLine("	CROSS APPLY Embalaje.nodes('//elementos/elemento') AS T(C)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("-- Transporte Adicional")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpFinal(	Documento, Fec_Ini, Fec_Fin, Cod_Cli, Nom_Cli, Status,")
            loComandoSeleccionar.AppendLine("						Mon_Net, For_Env, Tip_Env, Dir_Ent, Grupo,")
            loComandoSeleccionar.AppendLine("						Ta_Renglon, Ta_Transporte, Ta_Referencia, Ta_Inicio,")
            loComandoSeleccionar.AppendLine("						Ta_Hor_Ini, Ta_Cantidad, Ta_For_Env, Ta_Operador)")
            loComandoSeleccionar.AppendLine("SELECT	#tmpPedidos.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Fec_Ini									AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Fec_Fin									AS Fec_Fin,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Cod_Cli									AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Nom_Cli									AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Status									AS Status,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Mon_Net									AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.For_Env									AS For_Env, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Tip_Env									AS Tip_Env, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Dir_Ent									AS Dir_Ent,")
            loComandoSeleccionar.AppendLine("		CAST('TA' AS CHAR(2))								AS Grupo,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@renglon)[1]',		'INT')				AS Ta_Renglon,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@cod_tra)[1]',		'VARCHAR(MAX)')		AS Ta_Transporte,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@referencia)[1]',	'VARCHAR(MAX)')		AS Ta_Referencia,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@inicio)[1]',		'DATETIME')			AS Ta_Inicio,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@hor_ini)[1]',		'VARCHAR(MAX)')		AS Ta_Hor_Ini,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@can_uni)[1]',		'DECIMAL(28,10)')	AS Ta_Cantidad,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@for_env)[1]',		'VARCHAR(MAX)')		AS Ta_For_Env,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@operador)[1]', 	'VARCHAR(MAX)')		AS Ta_Operador")
            loComandoSeleccionar.AppendLine("FROM	#tmpPedidos")
            loComandoSeleccionar.AppendLine("	CROSS APPLY Tra_Adi.nodes('//elementos/elemento') AS T(C)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("-- Seguimientos Logísticos")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpFinal(	Documento, Fec_Ini, Fec_Fin, Cod_Cli, Nom_Cli, Status,")
            loComandoSeleccionar.AppendLine("						Mon_Net, For_Env, Tip_Env, Dir_Ent, Grupo,")
            loComandoSeleccionar.AppendLine("						Sl_Renglon, Sl_Fecha, Sl_Hora, Sl_Contacto, Sl_Accion, ")
            loComandoSeleccionar.AppendLine("						Sl_Medio, Sl_Comentario, Sl_Prioridad)")
            loComandoSeleccionar.AppendLine("SELECT	#tmpPedidos.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Fec_Ini									AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Fec_Fin									AS Fec_Fin,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Cod_Cli									AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Nom_Cli									AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Status									AS Status,")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Mon_Net									AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.For_Env									AS For_Env, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Tip_Env									AS Tip_Env, ")
            loComandoSeleccionar.AppendLine("		#tmpPedidos.Dir_Ent									AS Dir_Ent,	")
            loComandoSeleccionar.AppendLine("		CAST('SL' AS CHAR(2))								AS Grupo,	")
            loComandoSeleccionar.AppendLine("		ROW_NUMBER()OVER(												")
            loComandoSeleccionar.AppendLine("			PARTITION BY #tmpPedidos.Documento							")
            loComandoSeleccionar.AppendLine("			ORDER BY T.C.value('(@fecha)[1]', 'DATETIME'))	AS Sl_Renglon,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@fecha)[1]',		'DATETIME')			AS Sl_Fecha,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@hora)[1]',			'VARCHAR(10)')		AS Sl_Hora,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@contacto)[1]',		'VARCHAR(MAX)')		AS Sl_Contacto,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@accion)[1]',		'VARCHAR(MAX)')		AS Sl_Accion,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@medio)[1]',		'VARCHAR(MAX)')		AS Sl_Medio,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@comentario)[1]',	'VARCHAR(MAX)')		AS Sl_Comentario,")
            loComandoSeleccionar.AppendLine("		T.C.value('(@prioridad)[1]',	'VARCHAR(MAX)')		AS Sl_Prioridad")
            loComandoSeleccionar.AppendLine("FROM	#tmpPedidos")
            loComandoSeleccionar.AppendLine("	CROSS APPLY Seg_Log.nodes('//elementos/elemento') AS T(C)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("-- Documentaciones")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpFinal(	Documento, Fec_Ini, Fec_Fin, Cod_Cli, Nom_Cli, Status,")
            loComandoSeleccionar.AppendLine("						Mon_Net, For_Env, Tip_Env, Dir_Ent, Grupo,")
            loComandoSeleccionar.AppendLine("						Do_Renglon, Do_Nom_Doc, Do_Archivo, Do_Ruta)")
            loComandoSeleccionar.AppendLine("SELECT		#tmpPedidos.Documento						AS Documento,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Fec_Ini							AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Fec_Fin							AS Fec_Fin,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Cod_Cli							AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Nom_Cli							AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Status							AS Status,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Mon_Net							AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.For_Env							AS For_Env, ")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Tip_Env							AS Tip_Env, ")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Dir_Ent							AS Dir_Ent,	")
            loComandoSeleccionar.AppendLine("			CAST('DO' AS CHAR(2))						AS Grupo,	")
            loComandoSeleccionar.AppendLine("			ROW_NUMBER() OVER(											")
            loComandoSeleccionar.AppendLine("				PARTITION BY #tmpPedidos.Documento						")
            loComandoSeleccionar.AppendLine("				ORDER BY Documentacion.Nom_Doc)			AS Do_Renglon,	")
            loComandoSeleccionar.AppendLine("			Documentacion.Nom_Doc						AS Do_Nom_Doc,	")
            loComandoSeleccionar.AppendLine("			Documentacion.Archivo						AS Do_Archivo,	")
            loComandoSeleccionar.AppendLine("			Documentacion.Ruta							AS Do_Ruta")
            loComandoSeleccionar.AppendLine("FROM		#tmpPedidos")
            loComandoSeleccionar.AppendLine("	JOIN	Documentacion ")
            loComandoSeleccionar.AppendLine("		ON	Documentacion.Cod_Reg = #tmpPedidos.Documento")
            loComandoSeleccionar.AppendLine("		AND	Documentacion.Origen = 'Pedidos'")		   
            loComandoSeleccionar.AppendLine("ORDER BY	#tmpPedidos.Documento, Documentacion.Nom_Doc")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("-- Resto de las Pedidos (las que no tienen Detalle XML de Logística)")
            loComandoSeleccionar.AppendLine("-- *********************************************************")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpFinal(	Documento, Fec_Ini, Fec_Fin, Cod_Cli, Nom_Cli, Status,")
            loComandoSeleccionar.AppendLine("						Mon_Net, For_Env, Tip_Env, Dir_Ent)")
            loComandoSeleccionar.AppendLine("SELECT		#tmpPedidos.Documento						AS Documento,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Fec_Ini							AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Fec_Fin							AS Fec_Fin,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Cod_Cli							AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Nom_Cli							AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Status							AS Status,")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Mon_Net							AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.For_Env							AS For_Env, ")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Tip_Env							AS Tip_Env, ")
            loComandoSeleccionar.AppendLine("			#tmpPedidos.Dir_Ent							AS Dir_Ent")
            loComandoSeleccionar.AppendLine("FROM		#tmpPedidos")
            loComandoSeleccionar.AppendLine("WHERE NOT EXISTS(SELECT * FROM #tmpFinal WHERE #tmpFinal.Documento = #tmpPedidos.Documento)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpPedidos")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		-- Encabezado")
            loComandoSeleccionar.AppendLine("			Documento, Fec_Ini, Fec_Fin, Cod_Cli, Nom_Cli, Status, Mon_Net, ")
            loComandoSeleccionar.AppendLine("			-- Sipping (también desde encabezado)")
            loComandoSeleccionar.AppendLine("			For_Env, Tip_Env, Dir_Ent, Grupo,")
            loComandoSeleccionar.AppendLine("			-- Grupo de Tiempos Logísticos")
            loComandoSeleccionar.AppendLine("			Tl_Renglon, Tl_Actividad, Tl_Inicio, Tl_Hor_Ini, Tl_Comentario, Tl_Por_Eje,")
            loComandoSeleccionar.AppendLine("			-- Grupo de Tiempos Logísticos")
            loComandoSeleccionar.AppendLine("			Em_Renglon, Em_Tipo, Em_Cantidad, Em_Peso, Em_Unidad, Em_Comentario,")
            loComandoSeleccionar.AppendLine("			-- Grupo de Tiempos Logísticos")
            loComandoSeleccionar.AppendLine("			Ta_Renglon, Ta_Transporte, Ta_Referencia, Ta_Inicio, ")
            loComandoSeleccionar.AppendLine("			Ta_Hor_Ini, Ta_Cantidad, Ta_For_Env, Ta_Operador,")
            loComandoSeleccionar.AppendLine("			-- Grupo de Seguimientos Logísticos")
            loComandoSeleccionar.AppendLine("			Sl_Renglon, Sl_Fecha, Sl_Hora, Sl_Contacto, Sl_Accion, Sl_Medio, Sl_Comentario, Sl_Prioridad,")
            loComandoSeleccionar.AppendLine("			-- Grupo de Documentación")
            loComandoSeleccionar.AppendLine("			Do_Renglon, Do_Nom_Doc, Do_Archivo,")  
            
            Dim lcRuta As String = goServicios.mObtenerCampoFormatoSQL(Me.mObtenerPathAbsoluto())
            
            
            loComandoSeleccionar.AppendLine("			REPLACE(Do_Ruta, '../../', " & lcRuta & ") AS  Do_Ruta")
            loComandoSeleccionar.AppendLine("FROM		#tmpFinal")
            'loComandoSeleccionar.AppendLine("ORDER BY	Documento, Grupo")
            loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE	#tmpFinal")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

											  

            Dim loServicios As New cusDatos.goDatos
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
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


            loObjetoReporte	=  cusAplicacion.goReportes.mCargarReporte("rSeguimientos_lPedidos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrSeguimientos_lPedidos.ReportSource =	 loObjetoReporte	


        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try
	End Sub

	Private Function mObtenerPathAbsoluto() As String

		Dim loUrlPeticion As Uri = Me.Request.Url
		Dim loConstructor As UriBuilder = New UriBuilder(loUrlPeticion.Scheme, loUrlPeticion.Host, loUrlPeticion.Port)

		loConstructor.Path = VirtualPathUtility.ToAbsolute("~/")
		
		Return loConstructor.Uri.ToString()

	End Function

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
' RJG: 07/11/12: Codigo inicial
'-------------------------------------------------------------------------------------------'
