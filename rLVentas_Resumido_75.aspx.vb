'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLVentas_Resumido_75"
'-------------------------------------------------------------------------------------------'
Partial Class rLVentas_Resumido_75
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            
            Dim lcParametro3Desde As String = cusAplicacion.goReportes.paParametrosIniciales(3) 
            
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
		
        '************************************************************************************
		' Variable para usar un 0 "Decimal" en lugar de uno entero.							*
        '************************************************************************************
            loComandoSeleccionar.AppendLine("DECLARE @lnCero AS Decimal(28,10)")
            loComandoSeleccionar.AppendLine("SET	 @lnCero = 0;")
            loComandoSeleccionar.AppendLine("			")
        '************************************************************************************
        ' Variable para usar un 0 "Decimal" en lugar de uno entero.							*
        '************************************************************************************
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaDesde AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaHasta AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE @lcSucursalDesde AS CHAR(10)")
            loComandoSeleccionar.AppendLine("DECLARE @lcSucursalHasta AS CHAR(10)")
            loComandoSeleccionar.AppendLine("DECLARE @lcRevisionDesde AS CHAR(10)")
            loComandoSeleccionar.AppendLine("DECLARE @lcRevisionHasta AS CHAR(10)")
            loComandoSeleccionar.AppendLine("SET	 @ldFechaDesde = " & lcParametro0Desde & ";")
            loComandoSeleccionar.AppendLine("SET	 @ldFechaHasta = " & lcParametro0Hasta & ";")
            loComandoSeleccionar.AppendLine("SET	 @lcSucursalDesde = " & lcParametro1Desde & ";")
            loComandoSeleccionar.AppendLine("SET	 @lcSucursalHasta = " & lcParametro1Hasta & ";")
            loComandoSeleccionar.AppendLine("SET	 @lcRevisionDesde = " & lcParametro2Desde & ";")
            loComandoSeleccionar.AppendLine("SET	 @lcRevisionHasta = " & lcParametro2Hasta & ";")
            loComandoSeleccionar.AppendLine("			")

        '************************************************************************************
        ' Obtiene el detalle de los datos a mostrar y aplica los filtros del reporte		*
        '************************************************************************************
            loComandoSeleccionar.AppendLine("SELECT		CONVERT(NCHAR(10), Cuentas_Cobrar.Fec_Ini, 103)								AS Fecha1,				")
            loComandoSeleccionar.AppendLine("       	CONVERT(NCHAR(25), Cuentas_Cobrar.Fec_Ini, 121)								AS Fecha2,				")
            loComandoSeleccionar.AppendLine("       	DATEPART(WEEKDAY, Cuentas_Cobrar.Fec_Ini)									AS Dia,					")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Fiscal2														AS Documento,			")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Documento													AS Numero_Documento,	")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Doc_Ori														AS Documento_Origen,	")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Tip_Ori														AS Tipo_Origen,			")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Cla_Ori														AS Clase_Origen,		")
            loComandoSeleccionar.AppendLine("       	N'          '																AS Documento_Afectado,	")
            loComandoSeleccionar.AppendLine("       	N'          '																AS Comprobante_Retencion,")
            loComandoSeleccionar.AppendLine("       	@lnCero																		AS Monto_Retenido,		")
            loComandoSeleccionar.AppendLine("       	@lnCero																		AS Ventas_No_Sujetas,	")
            loComandoSeleccionar.AppendLine("       	@lnCero																		AS Exportacion,			")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Cod_Tip														AS Cod_Tip,				")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Cod_Suc														AS Cod_Suc,				")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Usu_Cre														AS Usu_Cre,				")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Fiscal1														AS Impresora,			")
            loComandoSeleccionar.AppendLine("       	ISNULL(Cajas.Cod_Caj, '')													AS Cod_Caj,				")
            loComandoSeleccionar.AppendLine("       	Clientes.Fiscal																AS Fiscal,				")
            loComandoSeleccionar.AppendLine("       	Clientes.Rif																AS Rif,					")
            loComandoSeleccionar.AppendLine("       	Clientes.Nom_Cli															AS Nom_Cli,				")
            loComandoSeleccionar.AppendLine("       	(CASE WHEN (Cuentas_Cobrar.Cod_Tip IN ('N/CR', 'RETIVA')) THEN -1 ELSE 1 END)	AS Signo,			")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Mon_Bru														AS Bruto,				")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Mon_Exe														AS Exento,				")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Por_Des														AS Por_Des,				")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Mon_Des														AS Mon_Des,				")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Por_Rec														AS Por_Rec,				")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Mon_Rec														AS Mon_Rec,				")
            loComandoSeleccionar.AppendLine("       	(Cuentas_Cobrar.Mon_Otr1+Cuentas_Cobrar.Mon_Otr2+Cuentas_Cobrar.Mon_Otr3)	AS Mon_Otr,				")
            loComandoSeleccionar.AppendLine("       	CAST(Cuentas_Cobrar.Dis_Imp AS XML)											AS Dis_Imp				")
            loComandoSeleccionar.AppendLine("INTO		#tmpRegistros ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cajas ON Cajas.Caracter1 = Cuentas_Cobrar.Fiscal1")
            loComandoSeleccionar.AppendLine("					AND Cajas.Caracter1 > ''")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Cobrar.Fec_Ini    BETWEEN @ldFechaDesde ")
            loComandoSeleccionar.AppendLine("			AND @ldFechaHasta ")
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Cod_Tip IN ('FACT','N/CR','RETIVA')")
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("       	AND Cuentas_Cobrar.Cod_Suc    BETWEEN @lcSucursalDesde ")
            loComandoSeleccionar.AppendLine("       	AND @lcSucursalHasta ")
            
            If lcParametro3Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN @lcRevisionDesde ")
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN @lcRevisionDesde")
            End If
            loComandoSeleccionar.AppendLine(" 			AND @lcRevisionHasta ")
            loComandoSeleccionar.AppendLine("			")
													    
        '************************************************************************************
        ' Revisa la distibución de impuestos por documento y los separa del XML				*
        '************************************************************************************
            loComandoSeleccionar.AppendLine("SELECT	Documento,	Cod_Tip,														")
            loComandoSeleccionar.AppendLine("		Fecha1,		Fecha2,		Dia,												")
            loComandoSeleccionar.AppendLine("		Cod_Suc,	Usu_Cre,														")
            loComandoSeleccionar.AppendLine("		Signo,		Fiscal,		Cod_Caj,											")
            loComandoSeleccionar.AppendLine("		Rif, Nom_cli,																")
            loComandoSeleccionar.AppendLine("		Documento_Origen, Tipo_Origen, Clase_Origen,								")
            loComandoSeleccionar.AppendLine("		Numero_Documento,	 														")
            loComandoSeleccionar.AppendLine("		Documento_Afectado, 														")
            loComandoSeleccionar.AppendLine("		Comprobante_Retencion,														")
            loComandoSeleccionar.AppendLine("		Monto_Retenido, 															")
            loComandoSeleccionar.AppendLine("		Ventas_No_Sujetas, 															")
            loComandoSeleccionar.AppendLine("		Exportacion,																")
            loComandoSeleccionar.AppendLine("		(CASE WHEN (Impresora='') THEN 'SIF' ELSE Impresora END)	AS Impresora,	")
            loComandoSeleccionar.AppendLine("		(Por_Des*Signo)												AS Por_Des,		")
            loComandoSeleccionar.AppendLine("		(Mon_Des*Signo)												AS Mon_Des,		")
            loComandoSeleccionar.AppendLine("		(Por_Rec*Signo)												AS Por_Rec,		")
            loComandoSeleccionar.AppendLine("		(Mon_Rec*Signo)												AS Mon_Rec,		")
            loComandoSeleccionar.AppendLine("		(Mon_Otr*Signo)												AS Mon_Otr,		")
            loComandoSeleccionar.AppendLine("		(Exento*Signo)												AS Exento,		")
            loComandoSeleccionar.AppendLine("		REPLACE(REPLACE(REPLACE(T.C.value('(./codigo)[1]', 'NVARCHAR(10)'), CHAR(10),''), CHAR(13),''), CHAR(9),'')									AS Codigo,	")
            loComandoSeleccionar.AppendLine("		CAST(REPLACE(REPLACE(REPLACE(T.C.value('(./porcentaje)[1]',	'NVARCHAR(20)')	, CHAR(10),''), CHAR(13),''), CHAR(9),'') AS DECIMAL(28,10))	AS Porcentaje,	")
            loComandoSeleccionar.AppendLine("		CAST(REPLACE(REPLACE(REPLACE(T.C.value('(./base)[1]',	'NVARCHAR(20)')	, CHAR(10),''), CHAR(13),''), CHAR(9),'') AS DECIMAL(28,10))*Signo	AS Base,		")
            loComandoSeleccionar.AppendLine("		CAST(REPLACE(REPLACE(REPLACE(T.C.value('(./monto)[1]',	'NVARCHAR(20)')	, CHAR(10),''), CHAR(13),''), CHAR(9),'') AS DECIMAL(28,10))*Signo	AS Impuesto		")
            loComandoSeleccionar.AppendLine("INTO	#tmpImpuestoDistribuido")
            loComandoSeleccionar.AppendLine("FROM	#tmpRegistros")
            loComandoSeleccionar.AppendLine("	CROSS APPLY #tmpRegistros.Dis_Imp.nodes('/impuestos/impuesto') AS T(C)")
            loComandoSeleccionar.AppendLine("WHERE	CAST(REPLACE(REPLACE(REPLACE(T.C.value('(./porcentaje)[1]',	'NVARCHAR(20)')	, CHAR(10),''), CHAR(13),''), CHAR(9),'') AS DECIMAL(28,10)) > @lnCero")
            loComandoSeleccionar.AppendLine("	OR	Bruto = Exento	")
            
        '************************************************************************************
        ' La primera tabla temporal ya no es necesaria.										*
        '************************************************************************************
            loComandoSeleccionar.AppendLine("			")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpRegistros")
            loComandoSeleccionar.AppendLine("			")
            
        '************************************************************************************
        ' Busca la información de la factura de origen de las devoluciones y retenciones.	*
        '************************************************************************************
			loComandoSeleccionar.AppendLine("SELECT	F.Fiscal2						AS Fis_Afectado, 		")
			loComandoSeleccionar.AppendLine("		F.Documento						AS Doc_Afectado, 		")
			loComandoSeleccionar.AppendLine("		ISNULL(R.Num_Com, '')			AS Numero_Comprobante, 	")
			loComandoSeleccionar.AppendLine("		ID.Documento					AS Documento,			")
			loComandoSeleccionar.AppendLine("		ID.Numero_Documento				AS Numero_Documento,	")
			loComandoSeleccionar.AppendLine("		ID.Cod_Tip						AS Tipo_Documento,		")
			loComandoSeleccionar.AppendLine("		ID.Tipo_Origen					AS Tipo_Origen ")
			loComandoSeleccionar.AppendLine("INTO	#tmpDevoluciones										")
			loComandoSeleccionar.AppendLine("FROM	Cuentas_Cobrar AS Origen								")
			loComandoSeleccionar.AppendLine("	JOIN #tmpImpuestoDistribuido AS ID							")
			loComandoSeleccionar.AppendLine("		ON ID.Documento_Origen = Origen.Documento				")
			loComandoSeleccionar.AppendLine("		AND ID.Cod_Tip IN ('RETIVA')							")
			loComandoSeleccionar.AppendLine("		AND Origen.Cod_Tip = 'RETIVA'							")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Retenciones_Documentos AS R						")
			loComandoSeleccionar.AppendLine("		ON	R.Doc_Des = ID.Numero_Documento						")
			loComandoSeleccionar.AppendLine("		AND	R.Cla_Des = 'RETIVA'								")
			loComandoSeleccionar.AppendLine("		AND R.Tip_Des = 'Cuentas_Cobrar'						")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Cobrar AS F								")
			loComandoSeleccionar.AppendLine("		ON	F.Documento = R.Doc_Ori								")
			loComandoSeleccionar.AppendLine("		AND F.Cod_Tip = 'FACT'")
			loComandoSeleccionar.AppendLine("		AND R.Tip_Ori = 'Cuentas_Cobrar'")
			loComandoSeleccionar.AppendLine("		AND R.Cla_Ori = 'FACT'")
			loComandoSeleccionar.AppendLine("UNION ALL 																	")
 			loComandoSeleccionar.AppendLine("SELECT	Origen.Fiscal2					AS Fis_Afectado, 					")
			loComandoSeleccionar.AppendLine("		Origen.Documento				AS Doc_Afectado, 					")
			loComandoSeleccionar.AppendLine("		''					AS Numero_Comprobante, 							")
			loComandoSeleccionar.AppendLine("		ID.Documento					AS Documento,						")
			loComandoSeleccionar.AppendLine("		ID.Numero_Documento				AS Numero_Documento,				")
			loComandoSeleccionar.AppendLine("		ID.Cod_Tip						AS Tipo_Documento,					")
			loComandoSeleccionar.AppendLine("		ID.Tipo_Origen					AS Tipo_Origen						")
			loComandoSeleccionar.AppendLine("FROM	#tmpImpuestoDistribuido AS ID										")
			loComandoSeleccionar.AppendLine("	JOIN Renglones_dClientes AS RC											")
			loComandoSeleccionar.AppendLine("		ON RC.Documento = ID.Documento_Origen								")
			loComandoSeleccionar.AppendLine("		AND ID.Cod_Tip IN ('N/CR')											")
			loComandoSeleccionar.AppendLine("		AND RC.Renglon = 1													")
			loComandoSeleccionar.AppendLine("		AND ID.Tipo_Origen = 'Devoluciones_Clientes'						")
			loComandoSeleccionar.AppendLine("	JOIN Cuentas_Cobrar AS Origen											")
			loComandoSeleccionar.AppendLine("		ON Origen.Cod_Tip = 'FACT' 											")
			loComandoSeleccionar.AppendLine("		AND Origen.Documento = RC.Doc_Ori									")
			loComandoSeleccionar.AppendLine("		AND RC.Renglon = 1													")
			loComandoSeleccionar.AppendLine("UNION ALL 																	")
			loComandoSeleccionar.AppendLine("SELECT	Origen.Fiscal2					AS Fis_Afectado, 					")
			loComandoSeleccionar.AppendLine("		Origen.Documento				AS Doc_Afectado, 					")
			loComandoSeleccionar.AppendLine("		''					AS Numero_Comprobante, 							")
			loComandoSeleccionar.AppendLine("		ID.Documento					AS Documento,						")
			loComandoSeleccionar.AppendLine("		ID.Numero_Documento				AS Numero_Documento,				")
			loComandoSeleccionar.AppendLine("		ID.Cod_Tip						AS Tipo_Documento,					")
			loComandoSeleccionar.AppendLine("		ID.Tipo_Origen					AS Tipo_Origen						")
			loComandoSeleccionar.AppendLine("FROM	#tmpImpuestoDistribuido AS ID										")
			loComandoSeleccionar.AppendLine("	JOIN Cuentas_Cobrar AS Origen											")
			loComandoSeleccionar.AppendLine("		ON Origen.Cod_Tip = 'FACT' 											")
			loComandoSeleccionar.AppendLine("		AND Origen.Documento = ID.Documento_Origen							")
			loComandoSeleccionar.AppendLine("		AND ID.Tipo_Origen = 'Facturas'										")
			loComandoSeleccionar.AppendLine("			")
			loComandoSeleccionar.AppendLine("			")
			loComandoSeleccionar.AppendLine("			")
			loComandoSeleccionar.AppendLine("			")
			loComandoSeleccionar.AppendLine("			")
			
        '************************************************************************************
        ' Crea la tabla de sustitución para los días de semana.								*
        '************************************************************************************
            loComandoSeleccionar.AppendLine("DECLARE @tmpDias AS TABLE(Dia INT, Dia_Sem NVARCHAR(20))")
            loComandoSeleccionar.AppendLine("INSERT INTO @tmpDias VALUES (1, 'Domingo')")
            loComandoSeleccionar.AppendLine("INSERT INTO @tmpDias VALUES (2, 'Lunes')")
            loComandoSeleccionar.AppendLine("INSERT INTO @tmpDias VALUES (3, 'Martes')")
            loComandoSeleccionar.AppendLine("INSERT INTO @tmpDias VALUES (4, 'Miércoles')")
            loComandoSeleccionar.AppendLine("INSERT INTO @tmpDias VALUES (5, 'Jueves')")
            loComandoSeleccionar.AppendLine("INSERT INTO @tmpDias VALUES (6, 'Viernes')")
            loComandoSeleccionar.AppendLine("INSERT INTO @tmpDias VALUES (7, 'Sábado')")
            loComandoSeleccionar.AppendLine("			")
            
			

        '************************************************************************************
        ' Selecciona los registros correspondientes a los "No Contribuyentes",				*
        ' Agrupados por día y caja (Solo FACT y RETIVA).									*
        '************************************************************************************
            loComandoSeleccionar.AppendLine("SELECT		'A'																		AS Tipo_Operacion,			")
            loComandoSeleccionar.AppendLine("			T.Fecha1																AS Fecha1,					")
            loComandoSeleccionar.AppendLine("			CAST(T.Fecha2 AS NCHAR(10))												AS Fecha2,					")
            loComandoSeleccionar.AppendLine("			Dias_Semana.Dia_Sem														AS Dia_Sem, 				")
            loComandoSeleccionar.AppendLine("			T.Cod_Suc																AS Cod_Suc, 				")
            loComandoSeleccionar.AppendLine("			T.Cod_Caj																AS Cod_Caj, 				")
            loComandoSeleccionar.AppendLine("			T.Impresora																AS Impresora,				")
            loComandoSeleccionar.AppendLine("			MIN(F.Documento) 														AS Doc_Ini, 				")
            loComandoSeleccionar.AppendLine("			MAX(F.Documento) 														AS Doc_Fin, 				")
            loComandoSeleccionar.AppendLine("			''				 														AS Doc_Ncr,                 ")
            loComandoSeleccionar.AppendLine("			SUM(T.Mon_Otr)															AS Mon_Otr, 				")
            loComandoSeleccionar.AppendLine("			''																		AS Numero_Documento,		")
            loComandoSeleccionar.AppendLine("			''																		AS Tipo_Documento, 			")
            loComandoSeleccionar.AppendLine("			''																		AS Documento_AfectadoNC,	")
            loComandoSeleccionar.AppendLine("			''																		AS Documento_Afectado, 		")
            loComandoSeleccionar.AppendLine("			''																		AS Comprobante_Retencion, 	")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Monto_Retenido, 			")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Ventas_No_Sujetas, 		")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Exportacion,				")
            loComandoSeleccionar.AppendLine("			''																		AS Rif,						")
            loComandoSeleccionar.AppendLine("			''																		AS Nom_Cli,					")
            loComandoSeleccionar.AppendLine("			MAX(T.Porcentaje)														AS Porcentaje,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_DesC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_RecC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_BruC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_ExeC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_ImpC,				")
            loComandoSeleccionar.AppendLine("			SUM(T.Por_Des*T.Base/100)												AS Mon_DesNC,				")
            loComandoSeleccionar.AppendLine("			SUM(T.Por_Rec*T.Base/100)												AS Mon_RecNC,				")
            loComandoSeleccionar.AppendLine("			SUM(CASE WHEN T.Impuesto  = 0 THEN @lnCero ELSE T.Base END)				AS Mon_BruNC,				")
            loComandoSeleccionar.AppendLine("			SUM(T.Exento)															AS Mon_ExeNC,				")
            loComandoSeleccionar.AppendLine("			SUM(T.Impuesto)															AS Mon_ImpNC, 				")
            loComandoSeleccionar.AppendLine("			SUM((CASE WHEN T.Cod_Tip = 'FACT' THEN T.Base		ELSE @lnCero END))	AS Fact_BI,					")
            loComandoSeleccionar.AppendLine("			SUM((CASE WHEN T.Cod_Tip = 'FACT' THEN T.Impuesto	ELSE @lnCero END))	AS Fact_IMP,				")
            loComandoSeleccionar.AppendLine("			SUM((CASE WHEN T.Cod_Tip = 'FACT' THEN T.Exento		ELSE @lnCero END))	AS Fact_EXE,				")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Fact_EXO,				")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Fact_NS,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Fact_SDCF,				")
            loComandoSeleccionar.AppendLine("			SUM(CASE WHEN T.Cod_Tip = 'RETIVA' THEN T.Exento	ELSE @lnCero END)	AS Fact_Retenido,			")
            loComandoSeleccionar.AppendLine("			SUM((CASE WHEN T.Cod_Tip = 'N/CR' THEN -(CASE WHEN T.Impuesto  = 0 THEN @lnCero ELSE T.Base END)		ELSE @lnCero END))	AS Ncr_BI,					")
            loComandoSeleccionar.AppendLine("			SUM((CASE WHEN T.Cod_Tip = 'N/CR' THEN -T.Impuesto	ELSE @lnCero END))	AS Ncr_IMP,					")
            loComandoSeleccionar.AppendLine("			SUM((CASE WHEN T.Cod_Tip = 'N/CR' THEN -T.Exento	ELSE @lnCero END))	AS Ncr_EXE,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Ncr_EXO,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Ncr_NS,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Ncr_SDCF					")
            loComandoSeleccionar.AppendLine("INTO		#tmpParteA  																						")
            loComandoSeleccionar.AppendLine("FROM		#tmpImpuestoDistribuido AS T																		")
            loComandoSeleccionar.AppendLine("	JOIN	@tmpDias AS Dias_Semana ON Dias_Semana.Dia = T.Dia													")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpImpuestoDistribuido AS F ON F.Fecha2 = T.Fecha2												")
            loComandoSeleccionar.AppendLine("		AND	F.Cod_Caj = T.Cod_Caj																				")
            loComandoSeleccionar.AppendLine("		AND	F.Impresora = T.Impresora																			")
            loComandoSeleccionar.AppendLine("		AND	F.Documento = T.Documento																			")
            loComandoSeleccionar.AppendLine("		AND	F.Numero_Documento = T.Numero_Documento																")
            loComandoSeleccionar.AppendLine("		AND	F.Cod_Tip	= 'FACT'																				")
            loComandoSeleccionar.AppendLine("		AND	T.Cod_Tip	= 'FACT'    																			")
            loComandoSeleccionar.AppendLine("		AND F.Fiscal = 0																						")
            loComandoSeleccionar.AppendLine("		AND T.Fiscal = 0																						")
            loComandoSeleccionar.AppendLine("WHERE		T.Fiscal = 0																						")
            loComandoSeleccionar.AppendLine("   AND		T.Cod_Tip <> 'N/CR'																					")
            loComandoSeleccionar.AppendLine("GROUP BY	T.Cod_Suc, T.Cod_Caj,																				")
            loComandoSeleccionar.AppendLine("			T.Fecha1, CAST(T.Fecha2 AS NCHAR(10)), T.Impresora,													")
            loComandoSeleccionar.AppendLine("			Dias_Semana.Dia_Sem																					")
            loComandoSeleccionar.AppendLine("           ")
            loComandoSeleccionar.AppendLine("UNION ALL  ")
            loComandoSeleccionar.AppendLine("           ")

        '************************************************************************************
        ' Selecciona los registros correspondientes a los "No Contribuyentes",				*
        ' Agrupados por día y caja (Solo N/CR) >> Se unen en #tmpParteA.	             	*
        '************************************************************************************
            loComandoSeleccionar.AppendLine("SELECT		'B'																		AS Tipo_Operacion,			")
            loComandoSeleccionar.AppendLine("			T.Fecha1																AS Fecha1,					")
            loComandoSeleccionar.AppendLine("			CAST(T.Fecha2 AS NCHAR(10))												AS Fecha2,					")
            loComandoSeleccionar.AppendLine("			Dias_Semana.Dia_Sem														AS Dia_Sem, 				")
            loComandoSeleccionar.AppendLine("			T.Cod_Suc																AS Cod_Suc, 				")
            loComandoSeleccionar.AppendLine("			T.Cod_Caj																AS Cod_Caj, 				")
            loComandoSeleccionar.AppendLine("			T.Impresora																AS Impresora,				")
            loComandoSeleccionar.AppendLine("			''				 														AS Doc_Ini, 				")
            loComandoSeleccionar.AppendLine("			''				 														AS Doc_Fin, 				")
            loComandoSeleccionar.AppendLine("			F.Documento		 														AS Doc_Ncr,                 ")
            loComandoSeleccionar.AppendLine("			T.Mon_Otr   															AS Mon_Otr, 				")
            loComandoSeleccionar.AppendLine("			''																		AS Numero_Documento,		")
            loComandoSeleccionar.AppendLine("			''																		AS Tipo_Documento, 			")
            loComandoSeleccionar.AppendLine("			ISNULL(CASE WHEN D.Fis_Afectado='' THEN D.Doc_Afectado ELSE D.Fis_Afectado END, '')	AS Documento_AfectadoNC,	")
            loComandoSeleccionar.AppendLine("			''																		AS Documento_Afectado, 		")
            loComandoSeleccionar.AppendLine("			''																		AS Comprobante_Retencion, 	")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Monto_Retenido, 			")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Ventas_No_Sujetas, 		")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Exportacion,				")
            loComandoSeleccionar.AppendLine("			''																		AS Rif,						")
            loComandoSeleccionar.AppendLine("			''																		AS Nom_Cli,					")
            loComandoSeleccionar.AppendLine("			T.Porcentaje    														AS Porcentaje,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_DesC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_RecC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_BruC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_ExeC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_ImpC,				")
            loComandoSeleccionar.AppendLine("			(T.Por_Des*T.Base/100)													AS Mon_DesNC,				")
            loComandoSeleccionar.AppendLine("			(T.Por_Rec*T.Base/100)													AS Mon_RecNC,				")
            loComandoSeleccionar.AppendLine("			(CASE  WHEN T.Impuesto =0 THEN @lnCero ELSE T.Base END)				    AS Mon_BruNC,				")
            loComandoSeleccionar.AppendLine("			T.Exento															    AS Mon_ExeNC,				")
            loComandoSeleccionar.AppendLine("			T.Impuesto															    AS Mon_ImpNC, 				")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'FACT' THEN T.Base		ELSE @lnCero END)    	AS Fact_BI,					")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'FACT' THEN T.Impuesto	ELSE @lnCero END)    	AS Fact_IMP,				")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'FACT' THEN T.Exento		ELSE @lnCero END)     	AS Fact_EXE,				")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Fact_EXO,				")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Fact_NS,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Fact_SDCF,				")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'RETIVA' THEN T.Exento	ELSE @lnCero END)	    AS Fact_Retenido,			")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'N/CR' THEN -(CASE  WHEN T.Impuesto =0 THEN @lnCero ELSE T.Base END)		ELSE @lnCero END)	    AS Ncr_BI,					")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'N/CR' THEN -T.Impuesto	ELSE @lnCero END)	    AS Ncr_IMP,					")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'N/CR' THEN -T.Exento	ELSE @lnCero END)	    AS Ncr_EXE,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Ncr_EXO,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Ncr_NS,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Ncr_SDCF					")
            'loComandoSeleccionar.AppendLine("INTO		#tmpParteB                                                                                          ")
            loComandoSeleccionar.AppendLine("FROM		#tmpImpuestoDistribuido AS T																		")
            loComandoSeleccionar.AppendLine("	JOIN	@tmpDias AS Dias_Semana ON Dias_Semana.Dia = T.Dia													")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpImpuestoDistribuido AS F ON F.Fecha2 = T.Fecha2												")
            loComandoSeleccionar.AppendLine("		AND	F.Cod_Caj = T.Cod_Caj																				")
            loComandoSeleccionar.AppendLine("		AND	F.Impresora = T.Impresora																			")
            loComandoSeleccionar.AppendLine("		AND	F.Documento = T.Documento																			")
            loComandoSeleccionar.AppendLine("		AND	F.Numero_Documento = T.Numero_Documento																")
            loComandoSeleccionar.AppendLine("		AND	F.Cod_Tip	= 'N/CR'																				")
            loComandoSeleccionar.AppendLine("		AND	T.Cod_Tip	= 'N/CR'    																			")
            loComandoSeleccionar.AppendLine("		AND F.Fiscal = 0																						")
            loComandoSeleccionar.AppendLine("		AND T.Fiscal = 0																						")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpDevoluciones AS D 																			")
            loComandoSeleccionar.AppendLine("		ON	D.Numero_Documento = T.Numero_Documento																")
            loComandoSeleccionar.AppendLine("		AND	D.Documento = T.Documento																			")
            loComandoSeleccionar.AppendLine("		AND	D.Tipo_Documento = T.Cod_Tip																		")
            loComandoSeleccionar.AppendLine("WHERE		T.Fiscal = 0																						")
            loComandoSeleccionar.AppendLine("   AND		T.Cod_Tip = 'N/CR'																					")
            loComandoSeleccionar.AppendLine("           ")

        '************************************************************************************
        ' Selecciona los registros correspondientes a los "Contribuyentes", sin agrupar.	*
        '************************************************************************************
            loComandoSeleccionar.AppendLine("SELECT		'C'																		AS Tipo_Operacion,			")
            loComandoSeleccionar.AppendLine("			T.Fecha1																AS Fecha1,					")
            loComandoSeleccionar.AppendLine("			CAST(T.Fecha2 AS NCHAR(10))												AS Fecha2,					")
            loComandoSeleccionar.AppendLine("			Dias_Semana.Dia_Sem														AS Dia_Sem, 				")
            loComandoSeleccionar.AppendLine("			T.Cod_Suc																AS Cod_Suc, 				")
            loComandoSeleccionar.AppendLine("			T.Cod_Caj																AS Cod_Caj, 				")
            loComandoSeleccionar.AppendLine("			T.Impresora																AS Impresora,				")
            loComandoSeleccionar.AppendLine("			''																		AS Doc_Ini, 				")
            loComandoSeleccionar.AppendLine("			''																		AS Doc_Fin, 				")
            loComandoSeleccionar.AppendLine("			''				 														AS Doc_Ncr, 				")
            loComandoSeleccionar.AppendLine("			T.Mon_Otr																AS Mon_Otr, 				")
			loComandoSeleccionar.AppendLine("			CASE WHEN Cod_Tip = 'RETIVA'																		")
			loComandoSeleccionar.AppendLine("				THEN																							")
			loComandoSeleccionar.AppendLine("					ISNULL(D.Numero_Comprobante, T.Numero_Documento)											")
			loComandoSeleccionar.AppendLine("				ELSE																							")
			loComandoSeleccionar.AppendLine("					(CASE WHEN T.Documento=''																	")
			loComandoSeleccionar.AppendLine("						THEN T.Numero_Documento																	")
			loComandoSeleccionar.AppendLine("						ELSE T.Documento																		")
			loComandoSeleccionar.AppendLine("					END)																						")
			loComandoSeleccionar.AppendLine("			END																		AS Numero_Documento,		")
			loComandoSeleccionar.AppendLine("			T.Cod_Tip																AS Tipo_Documento, 			")
            loComandoSeleccionar.AppendLine("			ISNULL(CASE WHEN D.Fis_Afectado='' THEN D.Doc_Afectado ELSE D.Fis_Afectado END, '')	AS Documento_Afectado, 	")
            loComandoSeleccionar.AppendLine("			ISNULL(D.Numero_Comprobante, '')										AS Comprobante_Retencion,	")
            loComandoSeleccionar.AppendLine("			T.Monto_Retenido														AS Monto_Retenido, 			")
            loComandoSeleccionar.AppendLine("			T.Ventas_No_Sujetas														AS Ventas_No_Sujetas,		")
            loComandoSeleccionar.AppendLine("			T.Exportacion															AS Exportacion,				")
			loComandoSeleccionar.AppendLine("			T.Rif																	AS Rif,						")
			loComandoSeleccionar.AppendLine("			T.Nom_Cli																AS Nom_Cli,					")
            loComandoSeleccionar.AppendLine("			T.Porcentaje															AS Porcentaje,				")
            loComandoSeleccionar.AppendLine("			T.Por_Des*T.Base/100    												AS Mon_DesC,				")
            loComandoSeleccionar.AppendLine("			T.Por_Rec*T.Base/100    												AS Mon_RecC,				")
            loComandoSeleccionar.AppendLine("			T.Base																	AS Mon_BruC,				")
            loComandoSeleccionar.AppendLine("			T.Exento																AS Mon_ExeC,				")
            loComandoSeleccionar.AppendLine("			T.Impuesto																AS Mon_ImpC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_DesNC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_RecNC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_BruNC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_ExeNC,				")
            loComandoSeleccionar.AppendLine("			@lnCero 																AS Mon_ImpNC, 				")	
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'FACT' THEN T.Base		ELSE @lnCero END)		AS Fact_BI,					")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'FACT' THEN T.Impuesto	ELSE @lnCero END)		AS Fact_IMP,				")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'FACT' THEN T.Exento		ELSE @lnCero END)		AS Fact_EXE,				")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Fact_EXO,				")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Fact_NS,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Fact_SDCF,				")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'RETIVA' THEN T.Exento	ELSE @lnCero END)		AS Fact_Retenido,			")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'N/CR' THEN T.Base		ELSE @lnCero END)		AS Ncr_BI,					")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'N/CR' THEN T.Impuesto	ELSE @lnCero END)		AS Ncr_IMP,					")
            loComandoSeleccionar.AppendLine("			(CASE WHEN T.Cod_Tip = 'N/CR' THEN T.Exento		ELSE @lnCero END)		AS Ncr_EXE,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Ncr_EXO,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Ncr_NS,					")
            loComandoSeleccionar.AppendLine("			@lnCero																	AS Ncr_SDCF					")
            loComandoSeleccionar.AppendLine("INTO		#tmpParteC                  																		")
            loComandoSeleccionar.AppendLine("FROM		#tmpImpuestoDistribuido AS T																		")
            loComandoSeleccionar.AppendLine("	JOIN	@tmpDias AS Dias_Semana ON Dias_Semana.Dia = T.Dia													")
            loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpDevoluciones AS D 																			")
            loComandoSeleccionar.AppendLine("		ON	D.Numero_Documento = T.Numero_Documento																")
            loComandoSeleccionar.AppendLine("		AND	D.Documento = T.Documento																			")
            loComandoSeleccionar.AppendLine("		AND	D.Tipo_Documento = T.Cod_Tip																		")
            loComandoSeleccionar.AppendLine("WHERE		T.Fiscal = 1																						")
            loComandoSeleccionar.AppendLine("ORDER BY	Fecha2 ASC, T.Cod_Caj ASC, Tipo_Operacion ASC, Numero_Documento										")
            loComandoSeleccionar.AppendLine("			")

        '************************************************************************************
        ' Borra las tablas remporales innecesarias.                                     	*
        '************************************************************************************
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpDevoluciones")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpImpuestoDistribuido")
            loComandoSeleccionar.AppendLine("			")

        '************************************************************************************
        ' Borra las tablas remporales innecesarias.                                     	*
        '************************************************************************************
            loComandoSeleccionar.AppendLine("SELECT		ISNULL(A.Tipo_Operacion, C.Tipo_Operacion)	                            AS Tipo_Operacion, 			")
            loComandoSeleccionar.AppendLine("			ISNULL(A.Fecha1, C.Fecha1)							                    AS Fecha1, 			        ")
            loComandoSeleccionar.AppendLine("			ISNULL(A.Fecha2, C.Fecha2)							                    AS Fecha2, 			        ")
            loComandoSeleccionar.AppendLine("			ISNULL(A.Cod_Suc, C.Cod_Suc)    						                AS Cod_Suc, 			    ")
            loComandoSeleccionar.AppendLine("			ISNULL(A.Cod_Caj, C.Cod_Caj)     						                AS Cod_Caj, 			    ")
            loComandoSeleccionar.AppendLine("			ISNULL(A.Impresora, C.Impresora) 			                    		AS Impresora, 			    ")
            loComandoSeleccionar.AppendLine("			ISNULL(A.Doc_ini, '')								 					AS Doc_ini, 			    ")
            loComandoSeleccionar.AppendLine("			ISNULL(A.Doc_Fin, '')								 					AS Doc_Fin, 			    ")
            loComandoSeleccionar.AppendLine("			ISNULL(A.Doc_Ncr, '')								 					AS Doc_Ncr, 			    ")
            loComandoSeleccionar.AppendLine("			ISNULL(A.Documento_AfectadoNC, '')										AS Documento_AfectadoNC, 	")
            loComandoSeleccionar.AppendLine("			C.Numero_Documento														AS Numero_Documento, 		")
            loComandoSeleccionar.AppendLine("			C.Tipo_Documento														AS Tipo_Documento, 			")
            loComandoSeleccionar.AppendLine("			C.Documento_Afectado													AS Documento_Afectado, 		")
            loComandoSeleccionar.AppendLine("			C.Comprobante_Retencion													AS Comprobante_Retencion,	")
            loComandoSeleccionar.AppendLine("			C.Monto_Retenido														AS Monto_Retenido, 			")
            loComandoSeleccionar.AppendLine("			C.Ventas_No_Sujetas														AS Ventas_No_Sujetas, 		")
            loComandoSeleccionar.AppendLine("			C.Exportacion															AS Exportacion, 			")
            loComandoSeleccionar.AppendLine("			C.Rif																	AS Rif, 			        ")
            loComandoSeleccionar.AppendLine("			C.Nom_cli																AS Nom_cli, 			    ")
            loComandoSeleccionar.AppendLine("			A.Porcentaje															AS Porcentaje,  			")
            loComandoSeleccionar.AppendLine("			C.Porcentaje															AS PorcentajeC, 			")
            loComandoSeleccionar.AppendLine("			A.Mon_DesNC										                        AS Mon_DesNC,  			    ")
            loComandoSeleccionar.AppendLine("			A.Mon_RecNC										                        AS Mon_RecNC,  			    ")
            loComandoSeleccionar.AppendLine("			A.Mon_BruNC										                        AS Mon_BruNC,   			")
            loComandoSeleccionar.AppendLine("			A.Mon_ExeNC										                        AS Mon_ExeNC,  			    ")
            loComandoSeleccionar.AppendLine("			A.Mon_ImpNC 										                    AS Mon_ImpNC,  			    ")
            loComandoSeleccionar.AppendLine("			C.Mon_DesC																AS Mon_DesC, 			    ")
            loComandoSeleccionar.AppendLine("			C.Mon_RecC																AS Mon_RecC, 			    ")
            loComandoSeleccionar.AppendLine("			C.Mon_BruC																AS Mon_BruC, 			    ")
            loComandoSeleccionar.AppendLine("			C.Mon_ExeC																AS Mon_ExeC, 			    ")
            loComandoSeleccionar.AppendLine("			C.Mon_ImpC																AS Mon_ImpC,  			    ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Fact_BI, @lnCero)		+ ISNULL(C.Fact_BI, @lnCero))			AS Fact_BI,                 ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Fact_IMP, @lnCero)	+ ISNULL(C.Fact_IMP, @lnCero))			AS Fact_IMP,                ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Fact_EXE, @lnCero)	+ ISNULL(C.Fact_EXE, @lnCero) +			                            ")
            loComandoSeleccionar.AppendLine("				ISNULL(A.Mon_Otr, @lnCero)	)										AS Fact_EXE,                ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Fact_Exo, @lnCero)	+ ISNULL(C.Fact_Exo, @lnCero))			AS Fact_Exo,                ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Fact_NS, @lnCero)		+ ISNULL(C.Fact_NS, @lnCero))			AS Fact_NS,                 ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Fact_SDCF, @lnCero)	+ ISNULL(C.Fact_SDCF, @lnCero))			AS Fact_SDCF,               ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Fact_Retenido, @lnCero) + ISNULL(C.Fact_Retenido, @lnCero))	AS Fact_Retenido,           ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Ncr_BI, @lnCero)		+ ISNULL(C.Ncr_BI, @lnCero))			AS Ncr_BI,                  ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Ncr_IMP, @lnCero)		+ ISNULL(C.Ncr_IMP, @lnCero))			AS Ncr_IMP,                 ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Ncr_EXE, @lnCero)		+ ISNULL(C.Ncr_EXE, @lnCero) +                                      ")
            loComandoSeleccionar.AppendLine("				ISNULL(A.Mon_Otr, @lnCero))											AS Ncr_EXE,                 ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Ncr_Exo, @lnCero)		+ ISNULL(C.Ncr_Exo, @lnCero))			AS Ncr_Exo,                 ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Ncr_NS, @lnCero)		+ ISNULL(C.Ncr_NS, @lnCero))			AS Ncr_NS,                  ")
            loComandoSeleccionar.AppendLine("			(ISNULL(A.Ncr_SDCF, @lnCero)	+ ISNULL(C.Ncr_SDCF, @lnCero))			AS Ncr_SDCF,                ")
            loComandoSeleccionar.AppendLine("           CASE WHEN (ISNULL(C.Tipo_Operacion, '') = '')                                                       ")
            loComandoSeleccionar.AppendLine("           	THEN 0                                                                                          ")
            loComandoSeleccionar.AppendLine("           	ELSE ROW_NUMBER()                                                                               ")
            loComandoSeleccionar.AppendLine("           		OVER (	PARTITION BY C.Fecha2, C.Cod_Caj, ISNULL(A.Tipo_Operacion, ISNULL(A.Tipo_Operacion, C.Tipo_Operacion))")
            loComandoSeleccionar.AppendLine("           				ORDER BY C.Fecha2 ASC, C.Cod_Caj ASC, ISNULL(A.Tipo_Operacion, ISNULL(A.Tipo_Operacion, C.Tipo_Operacion)) ASC,")
            loComandoSeleccionar.AppendLine("           						CASE WHEN (C.Tipo_Documento = 'FACT') THEN 0 ELSE 1 END ASC, C.Numero_Documento ASC) ")
            loComandoSeleccionar.AppendLine("           END                                                                     AS Repetido                 ")
            loComandoSeleccionar.AppendLine("INTO       #tmpFinal																							")
            loComandoSeleccionar.AppendLine("FROM		#tmpParteA AS A			                                                                            ")
            loComandoSeleccionar.AppendLine("	FULL JOIN #tmpParteC AS C			                                                                        ")
            loComandoSeleccionar.AppendLine("		ON	C.Fecha2 = A.Fecha2			                                                                        ")
            loComandoSeleccionar.AppendLine("		AND	C.Cod_Caj = A.Cod_Caj			                                                                    ")
            loComandoSeleccionar.AppendLine("		AND	C.Impresora = A.Impresora			                                                                ")
            loComandoSeleccionar.AppendLine("		AND A.Doc_Ini <> ''						                                                                ")
            loComandoSeleccionar.AppendLine("ORDER BY	Fecha2 ASC, Cod_Caj ASC, Tipo_Operacion ASC                                                         ")
            loComandoSeleccionar.AppendLine("			")

            loComandoSeleccionar.AppendLine("SELECT		Tipo_Operacion 				                    AS Tipo_Operacion, 		")
            loComandoSeleccionar.AppendLine("			Fecha1 						                    AS Fecha1, 			    ")
            loComandoSeleccionar.AppendLine("			Fecha2 						                    AS Fecha2, 			    ")
            loComandoSeleccionar.AppendLine("			Cod_Suc 					                    AS Cod_Suc, 			")
            loComandoSeleccionar.AppendLine("			Cod_Caj 					                    AS Cod_Caj, 			")
            loComandoSeleccionar.AppendLine("			Impresora 					                    AS Impresora, 			")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido <=1)	                    		                ")
            loComandoSeleccionar.AppendLine("				THEN	Doc_ini			                                            ")
            loComandoSeleccionar.AppendLine("				ELSE	NULL			                                            ")
            loComandoSeleccionar.AppendLine("			END		 					                    AS Doc_ini, 			")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido <=1)	                    		                ")
            loComandoSeleccionar.AppendLine("				THEN	Doc_Fin			                                            ")
            loComandoSeleccionar.AppendLine("				ELSE	NULL			                                            ")
            loComandoSeleccionar.AppendLine("			END		 					                    AS Doc_Fin, 			")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido <=1)	                    		                ")
            loComandoSeleccionar.AppendLine("				THEN	Doc_Ncr			                                            ")
            loComandoSeleccionar.AppendLine("				ELSE	NULL			                                            ")
            loComandoSeleccionar.AppendLine("			END		 					                    AS Doc_Ncr, 			")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido <=1)			                                    ")
            loComandoSeleccionar.AppendLine("				THEN	Documento_AfectadoNC			                            ")
            loComandoSeleccionar.AppendLine("				ELSE	NULL			                                            ")
            loComandoSeleccionar.AppendLine("			END		 					                    AS Documento_AfectadoNC,")
            loComandoSeleccionar.AppendLine("			Numero_Documento 			                    AS Numero_Documento, 	")
            loComandoSeleccionar.AppendLine("			Tipo_Documento 				                    AS Tipo_Documento, 		")
            loComandoSeleccionar.AppendLine("			Documento_Afectado 			                    AS Documento_Afectado, 	")
            loComandoSeleccionar.AppendLine("			Comprobante_Retencion		                    AS Comprobante_Retencion,")
            loComandoSeleccionar.AppendLine("			Monto_Retenido 				                    AS Monto_Retenido, 		")
            loComandoSeleccionar.AppendLine("			Ventas_No_Sujetas 			                    AS Ventas_No_Sujetas, 	")
            loComandoSeleccionar.AppendLine("			Exportacion 				                    AS Exportacion, 		")
            loComandoSeleccionar.AppendLine("			Rif 						                    AS Rif, 			    ")
            loComandoSeleccionar.AppendLine("			Nom_cli 					                    AS Nom_cli, 			")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido <=1 )	                                            ")
            loComandoSeleccionar.AppendLine("				THEN	Porcentaje		                                            ")
            loComandoSeleccionar.AppendLine("				ELSE	NULL			                                            ")
            loComandoSeleccionar.AppendLine("			END		 					AS Porcentaje, 			                    ")
            loComandoSeleccionar.AppendLine("			PorcentajeC 				AS PorcentajeC, 			                ")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido <=1)	                                        	")
            loComandoSeleccionar.AppendLine("				THEN	Mon_DesNC		                                        	")
            loComandoSeleccionar.AppendLine("				ELSE	NULL			                                            ")
            loComandoSeleccionar.AppendLine("			END		 					AS Mon_DesNC, 			                    ")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido <=1)	                                            ")
            loComandoSeleccionar.AppendLine("				THEN	Mon_RecNC		                                            ")
            loComandoSeleccionar.AppendLine("				ELSE	NULL			                                            ")
            loComandoSeleccionar.AppendLine("			END		 					AS Mon_RecNC, 			                    ")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido <=1)	                                            ")
            loComandoSeleccionar.AppendLine("				THEN	Mon_BruNC		                                            ")
            loComandoSeleccionar.AppendLine("				ELSE	NULL			                                            ")
            loComandoSeleccionar.AppendLine("			END		 					AS Mon_BruNC,  			                    ")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido <=1)	                                            ")
            loComandoSeleccionar.AppendLine("				THEN	Mon_ExeNC		                                            ")
            loComandoSeleccionar.AppendLine("				ELSE	NULL			                                            ")
            loComandoSeleccionar.AppendLine("			END		 					AS Mon_ExeNC, 			                    ")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido <=1)	                                            ")
            loComandoSeleccionar.AppendLine("				THEN	Mon_ImpNC		                                            ")
            loComandoSeleccionar.AppendLine("				ELSE	NULL			                                            ")
            loComandoSeleccionar.AppendLine("			END		 																AS Mon_ImpNC, 			                    ")
            loComandoSeleccionar.AppendLine("			Mon_DesC 																AS Mon_DesC, 			                    ")
            loComandoSeleccionar.AppendLine("			Mon_RecC 																AS Mon_RecC, 			                    ")
            loComandoSeleccionar.AppendLine("			Mon_BruC 																AS Mon_BruC, 			                    ")
            loComandoSeleccionar.AppendLine("			Mon_ExeC 																AS Mon_ExeC, 			                    ")
            loComandoSeleccionar.AppendLine("			Mon_ImpC 																AS Mon_ImpC, 			                    ")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido >1 AND Mon_BruC > 0)			                    ")
            loComandoSeleccionar.AppendLine("				THEN	Mon_BruC			                        ")
            loComandoSeleccionar.AppendLine("				ELSE	(CASE WHEN (Repetido<=1) THEN Fact_BI ELSE 0 END) ")
            loComandoSeleccionar.AppendLine("			END		 																AS Fact_BI, 			    ")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido >1 AND Mon_ImpC > 0) 			    ")
            loComandoSeleccionar.AppendLine("				THEN	Mon_ImpC			                        ")
            loComandoSeleccionar.AppendLine("				ELSE	(CASE WHEN (Repetido<=1) THEN Fact_IMP ELSE 0 END) ")
            loComandoSeleccionar.AppendLine("			END		 																AS Fact_IMP,  			    ")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido >1 AND Mon_ExeC > 0)			    ")
            loComandoSeleccionar.AppendLine("				THEN	Mon_ExeC			                        ")
            loComandoSeleccionar.AppendLine("				ELSE	(CASE WHEN (Repetido<=1) THEN Fact_EXE ELSE 0 END) ")
            loComandoSeleccionar.AppendLine("			END		 																	AS Fact_EXE, 			    ")
            loComandoSeleccionar.AppendLine("			Fact_Exo																	AS Fact_Exo, 			    ")
            loComandoSeleccionar.AppendLine("			Fact_NS 																	AS Fact_NS, 			    ")
            loComandoSeleccionar.AppendLine("			Fact_SDCF 																	AS Fact_SDCF, 			    ")
            loComandoSeleccionar.AppendLine("			Fact_Retenido																AS Fact_Retenido, 			")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido >1 AND Mon_BruC < 0) 			    ")
            loComandoSeleccionar.AppendLine("				THEN	Mon_BruC			                        ")
            loComandoSeleccionar.AppendLine("				ELSE	(CASE WHEN (Repetido<=1) THEN Ncr_BI ELSE 0 END)			")
            loComandoSeleccionar.AppendLine("			END		 																	AS Ncr_BI, 			        ")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido >1 AND Mon_ImpC < 0) 			    ")
            loComandoSeleccionar.AppendLine("				THEN	Mon_ImpC													")
            loComandoSeleccionar.AppendLine("				ELSE	(CASE WHEN (Repetido<=1) THEN Ncr_IMP ELSE 0 END) ")
            loComandoSeleccionar.AppendLine("			END		 																	AS Ncr_IMP,  			    ")
            loComandoSeleccionar.AppendLine("			CASE WHEN (Repetido >1 AND Mon_ExeC < 0) 			    				")
            loComandoSeleccionar.AppendLine("				THEN	Mon_ExeC			                        				")
            loComandoSeleccionar.AppendLine("				ELSE	(CASE WHEN (Repetido<=1) THEN Ncr_EXE ELSE 0 END) ")
            loComandoSeleccionar.AppendLine("			END		 																	AS Ncr_EXE, 			    ")
            loComandoSeleccionar.AppendLine("			Ncr_Exo 																	AS Ncr_Exo, 			    ")
            loComandoSeleccionar.AppendLine("			Ncr_NS 																		AS Ncr_NS, 			        ")
            loComandoSeleccionar.AppendLine("			Ncr_SDCF																	AS Ncr_SDCF,			    ")
            loComandoSeleccionar.AppendLine("			Repetido																	AS Repetido			        ")
            loComandoSeleccionar.AppendLine("FROM		#tmpFinal			                                    ")
            loComandoSeleccionar.AppendLine("ORDER BY	Fecha2 ASC, Cod_Caj ASC, Tipo_Operacion ASC, ")
            loComandoSeleccionar.AppendLine("			ISNULL(Doc_ini, 'zzzzzzzz' )ASC, ISNULL(Doc_Ncr, 'zzzzzzzz') ASC,")
            loComandoSeleccionar.AppendLine("			Repetido ASC")
            loComandoSeleccionar.AppendLine("			")

            Dim loServicios As New cusDatos.goDatos

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLVentas_Resumido_75", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLVentas_Resumido.ReportSource = loObjetoReporte

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
' RJG: 08/09/11: Codigo inicial (A partir de rLVentas_Resumido).							'
'-------------------------------------------------------------------------------------------'
' RJG: 13/09/11: Ajustes adicionales requeridos: columnas adicionales y cuadro resumen.		'
'-------------------------------------------------------------------------------------------'
' RJG: 20/09/11: Continuación de la programación.											'
'-------------------------------------------------------------------------------------------'
' RJG: 21/09/11: Continuación de la programación.											'
'-------------------------------------------------------------------------------------------'
' RJG: 19/10/11: Cierre de reporte con últimos cambios en presentación de N/CR detalladas y	'
'				 operaciones a Contribuyentes.												'
'-------------------------------------------------------------------------------------------'
' RJG: 02/11/11: Agregadas las N/CR relacionadas a Ordenes de Pago de Devoluciones de IPOS.	'
'-------------------------------------------------------------------------------------------'
' RJG: 07/10/13: Corrección en el cálculo del monto de Descuento: faltaba dividir el Por_Des'
'                entre 100.                                                                 '
'-------------------------------------------------------------------------------------------'
