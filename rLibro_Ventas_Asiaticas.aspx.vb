'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Ventas_GPV"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Ventas_GPV
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
            Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosFinales(5)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()
            loConsulta.AppendLine("CREATE TABLE #tmpTemporal(	Fecha				DATETIME, ")
            loConsulta.AppendLine("							Documento			VARCHAR(100),")
            loConsulta.AppendLine("							Cod_Tip				VARCHAR(100),")
            loConsulta.AppendLine("							Tipo				INT, ")
            loConsulta.AppendLine("							Tipo_Documento		VARCHAR(100),")
            loConsulta.AppendLine("							Reporte_Z			VARCHAR(100),  ")
            loConsulta.AppendLine("							Doc_Ini				VARCHAR(100), ")
            loConsulta.AppendLine("							Doc_fin				VARCHAR(100),")
            loConsulta.AppendLine("							status			VARCHAR(100),")
            loConsulta.AppendLine("							Cod_Cli				VARCHAR(100),")
            loConsulta.AppendLine("							Nom_Cli				VARCHAR(100),")
            loConsulta.AppendLine("							Rif					VARCHAR(100),")
            loConsulta.AppendLine("							Mon_Net				DECIMAL(28,10),")
            loConsulta.AppendLine("							Mon_Exe				DECIMAL(28,10),")
            loConsulta.AppendLine("							NC_Base				DECIMAL(28,10),")
            loConsulta.AppendLine("							NC_Por_Impuesto		DECIMAL(28,10),")
            loConsulta.AppendLine("							NC_Mon_Impuesto		DECIMAL(28,10),")
            loConsulta.AppendLine("							C_Base				DECIMAL(28,10),")
            loConsulta.AppendLine("							C_Por_Impuesto		DECIMAL(28,10),")
            loConsulta.AppendLine("							C_Mon_Impuesto		DECIMAL(28,10),")
            loConsulta.AppendLine("							Ret_Fecha			DATETIME,")
            loConsulta.AppendLine("							Documento_Afect		VARCHAR(100),")
            loConsulta.AppendLine("							Ret_Comprobante		VARCHAR(100),")
            loConsulta.AppendLine("							Ret_MontoRetenido	DECIMAL(28,10))")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- FACTURAS DE VENTA")
            loConsulta.AppendLine("INSERT INTO #tmpTemporal (	Fecha, Documento,Cod_Tip,Tipo, Tipo_Documento,Reporte_Z, Doc_Ini, Doc_fin,")
            loConsulta.AppendLine("							status,Cod_Cli,Nom_Cli,Rif,Mon_Net,Mon_Exe,NC_Base,NC_Por_Impuesto,")
            loConsulta.AppendLine("							NC_Mon_Impuesto,C_Base,C_Por_Impuesto,C_Mon_Impuesto,Ret_Fecha,Ret_Comprobante,")
            loConsulta.AppendLine("							Ret_MontoRetenido,Documento_Afect )")
            loConsulta.AppendLine("SELECT	CAST(Facturas.Fec_Ini AS DATE) 						AS Fecha, ")
            loConsulta.AppendLine("			Facturas.documento							        AS Documento, ")
            loConsulta.AppendLine("			'FACT'								                AS Cod_Tip, ")
            loConsulta.AppendLine("			1													AS Tipo, ")
            loConsulta.AppendLine("			'FACTURA'											AS Tipo_Documento, ")
            loConsulta.AppendLine("			Facturas.Referencia									AS Reporte_Z, ")
            loConsulta.AppendLine("			Facturas.Documento									AS Doc_Ini, ")
            loConsulta.AppendLine("			Facturas.Control									AS Doc_fin,")
            loConsulta.AppendLine("			Facturas.status							            AS status,")
            loConsulta.AppendLine("			Clientes.Cod_Cli									AS Cod_Cli,")
            loConsulta.AppendLine("			Clientes.Nom_Cli									AS Nom_Cli,")
            loConsulta.AppendLine("			Clientes.Rif										AS Rif,")
            loConsulta.AppendLine("			Facturas.Mon_Net									AS Mon_Net,")
            loConsulta.AppendLine("			Facturas.Mon_Exe									AS Mon_Exe,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 0 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Mon_Net - Facturas.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Base,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 0 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Por_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Por_Impuesto,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 0 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Mon_Impuesto,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Mon_Net - Facturas.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Base,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Por_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Por_Impuesto,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1 AND Facturas.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Facturas.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Mon_Impuesto,")
            loConsulta.AppendLine("			NULL                                                AS Ret_Fecha,")
            loConsulta.AppendLine("			''							                        AS Ret_Comprobante,")
            loConsulta.AppendLine("			0 							                        AS Ret_MontoRetenido,")
            loConsulta.AppendLine("			'' 							                        AS doc")
            loConsulta.AppendLine("FROM		Facturas")
            loConsulta.AppendLine("	JOIN	Renglones_Facturas ")
            loConsulta.AppendLine("		ON	Renglones_Facturas.Documento = Facturas.Documento")
            loConsulta.AppendLine("		AND Renglones_Facturas.Renglon = 1")
            loConsulta.AppendLine("	JOIN	Clientes ON Clientes.Cod_Cli = Facturas.Cod_Cli")
            loConsulta.AppendLine("WHERE    Facturas.Status <> 'Pendiente'")
            loConsulta.AppendLine("		AND Facturas.Mon_Net > 0")
            loConsulta.AppendLine("     AND Facturas.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("		AND Facturas.Documento    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("		AND Facturas.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("		AND Facturas.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("		AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Facturas.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Facturas.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)



            loConsulta.AppendLine("-- Facturas desde Cuentas por Cobrar")
            loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("SELECT		CAST(Cuentas_Cobrar.Fec_Ini AS DATE) 			AS Fecha, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.documento							AS Documento, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.cod_tip								AS Cod_Tip, ")
            loConsulta.AppendLine("			1													AS Tipo, ")
            loConsulta.AppendLine("			'FACTURA'									        AS Tipo_Documento,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Referencia							AS Reporte_Z,  ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Documento							AS Doc_Ini, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Control								AS Doc_fin,")
            loConsulta.AppendLine("			Cuentas_Cobrar.status						        AS status,")
            loConsulta.AppendLine("			Clientes.Cod_Cli									AS Cod_Cli,")
            loConsulta.AppendLine("			Clientes.Nom_Cli									AS Nom_Cli,")
            loConsulta.AppendLine("			Clientes.Rif										AS Rif,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Net								AS Mon_Net,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Exe								AS Mon_Exe,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 0 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Net - Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Base,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 0 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Por_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Por_Impuesto,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 0 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Mon_Impuesto,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Net - Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Base,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Por_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Por_Impuesto,")
            loConsulta.AppendLine("			(CASE WHEN Clientes.Fiscal = 1 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Mon_Impuesto,")
            loConsulta.AppendLine("			NULL                                                AS Ret_Fecha,")
            loConsulta.AppendLine("			''							                        AS Ret_Comprobante,")
            loConsulta.AppendLine("			0 							                        AS Ret_MontoRetenido,")
            loConsulta.AppendLine("			'' 							                        AS doc")
            loConsulta.AppendLine("FROM		Cuentas_Cobrar")
            loConsulta.AppendLine("	LEFT JOIN	Renglones_Documentos")
            loConsulta.AppendLine("		ON	Renglones_Documentos.Documento = Cuentas_Cobrar.Documento")
            loConsulta.AppendLine("		AND Renglones_Documentos.Cod_Tip = Cuentas_Cobrar.Cod_Tip")
            loConsulta.AppendLine("		AND Renglones_Documentos.Renglon = 1")
            loConsulta.AppendLine("		AND Renglones_Documentos.Origen = 'Ventas'")
            loConsulta.AppendLine("	JOIN	Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loConsulta.AppendLine("WHERE	Cuentas_Cobrar.Cod_Tip = 'FACT'")
            loConsulta.AppendLine("     AND Cuentas_Cobrar.automatico = 0")
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("		AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)

            loConsulta.AppendLine("-- NOTAS DE CRÉDITO")
            loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("SELECT		CAST(Cuentas_Cobrar.Fec_Ini AS DATE) 				AS Fecha, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.documento							AS Documento, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.cod_tip								AS Cod_Tip, ")
            loConsulta.AppendLine("			2													AS Tipo, ")
            loConsulta.AppendLine("			'N/CR'									AS Tipo_Documento,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Referencia							AS Reporte_Z,  ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Documento							AS Doc_Ini, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Control								AS Doc_fin,")
            loConsulta.AppendLine("			Renglones_Documentos.Cod_Alm						AS status,")
            loConsulta.AppendLine("			Clientes.Cod_Cli									AS Cod_Cli,")
            loConsulta.AppendLine("			Clientes.Nom_Cli									AS Nom_Cli,")
            loConsulta.AppendLine("			Clientes.Rif										AS Rif,")
            loConsulta.AppendLine("			-Cuentas_Cobrar.Mon_Net								AS Mon_Net,")
            loConsulta.AppendLine("			-Cuentas_Cobrar.Mon_Exe								AS Mon_Exe,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 0 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Net - Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Base,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 0 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Por_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Por_Impuesto,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 0 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS NC_Mon_Impuesto,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 1 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Net - Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Base,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 1 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Por_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Por_Impuesto,")
            loConsulta.AppendLine("			-(CASE WHEN Clientes.Fiscal = 1 AND Cuentas_Cobrar.Por_Imp1 > 0")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Imp1")
            loConsulta.AppendLine("				ELSE 0")
            loConsulta.AppendLine("			END)												AS C_Mon_Impuesto,")
            loConsulta.AppendLine("			NULL                                                AS Ret_Fecha,")
            loConsulta.AppendLine("			''   											    AS Ret_Comprobante,")
            loConsulta.AppendLine("			0 							                        AS Ret_MontoRetenido,")
            loConsulta.AppendLine("			CASE ")
            loConsulta.AppendLine("				WHEN Cuentas_cobrar.tip_ori = '' ")
            loConsulta.AppendLine("				THEN Cuentas_cobrar.factura")
            loConsulta.AppendLine("				WHEN Cuentas_cobrar.tip_ori = 'Cuentas_Cobrar' AND Cuentas_cobrar.cla_ori = 'FACT'")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.doc_ori")
            loConsulta.AppendLine("				ELSE 'DEV'")
            loConsulta.AppendLine("			END                                                 AS doc")
            loConsulta.AppendLine("FROM		Cuentas_Cobrar")
            loConsulta.AppendLine("	LEFT JOIN	Renglones_Documentos")
            loConsulta.AppendLine("		ON	Renglones_Documentos.Documento = Cuentas_Cobrar.Documento")
            loConsulta.AppendLine("		AND Renglones_Documentos.Cod_Tip = Cuentas_Cobrar.Cod_Tip")
            loConsulta.AppendLine("		AND Renglones_Documentos.Renglon = 1")
            loConsulta.AppendLine("		AND Renglones_Documentos.Origen = 'Ventas'")
            loConsulta.AppendLine("	JOIN	Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loConsulta.AppendLine("WHERE	Cuentas_Cobrar.Status IN ('Pendiente','Afectado','Pagado')")
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Tip = 'N/CR'")
            loConsulta.AppendLine("		AND ((Cuentas_Cobrar.tip_ori ='') OR ")
            loConsulta.AppendLine("			(Cuentas_cobrar.tip_ori = 'Cuentas_Cobrar' AND Cuentas_cobrar.cla_ori = 'FACT') OR")
            loConsulta.AppendLine("			(Cuentas_cobrar.tip_ori = 'devoluciones_clientes'))")
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("		AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE		#tmpTemporal")
            loConsulta.AppendLine("SET			Ret_Comprobante =/*	CAST(YEAR(Fecha) AS VARCHAR(4))+")
            loConsulta.AppendLine("								RIGHT('00'+CAST(MONTH(Fecha) AS VARCHAR(4)),2)+")
            loConsulta.AppendLine("								RIGHT('00'+CAST(DAY(Fecha) AS VARCHAR(4)),2)+*/")
            loConsulta.AppendLine("                             Retenciones.Com_Ret,")
            loConsulta.AppendLine("			Ret_MontoRetenido = Retenciones.mon_net,")
            loConsulta.AppendLine("			ret_fecha = Retenciones.fec_ini")
            loConsulta.AppendLine("FROM	(	SELECT		#tmpTemporal.Documento		AS Documento,")
            loConsulta.AppendLine("						#tmpTemporal.Cod_Tip			AS Cod_Tip,")
            loConsulta.AppendLine("						retenciones_documentos.num_com	AS Com_Ret,")
            loConsulta.AppendLine("						Cuentas_cobrar.mon_net,")
            loConsulta.AppendLine("						Cuentas_cobrar.fec_ini")
            loConsulta.AppendLine("            FROM retenciones_documentos")
            loConsulta.AppendLine("				JOIN	#tmpTemporal")
            loConsulta.AppendLine("					ON	#tmpTemporal.Documento = retenciones_documentos.doc_ori")
            loConsulta.AppendLine("					AND	#tmpTemporal.Cod_Tip = retenciones_documentos.cla_ori")
            loConsulta.AppendLine("					AND	retenciones_documentos.tip_ori = 'Cuentas_Cobrar'")
            loConsulta.AppendLine("					AND	retenciones_documentos.Clase = 'IMPUESTO'")
            loConsulta.AppendLine("            Join cuentas_cobrar")
            loConsulta.AppendLine("					ON	cuentas_cobrar.documento = retenciones_documentos.doc_des")
            loConsulta.AppendLine("					AND	cuentas_cobrar.cod_tip = retenciones_documentos.cla_des")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("		) AS Retenciones")
            loConsulta.AppendLine("WHERE	Retenciones.Documento = #tmpTemporal.documento")
            loConsulta.AppendLine("	AND	Retenciones.Cod_Tip = #tmpTemporal.Cod_Tip")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE #tmpTemporal")
            loConsulta.AppendLine("SET Mon_Net = 0, ")
            loConsulta.AppendLine("	Mon_Exe = 0,")
            loConsulta.AppendLine("	NC_BASE = 0,")
            loConsulta.AppendLine("	NC_Por_Impuesto = 0,")
            loConsulta.AppendLine("	NC_Mon_Impuesto = 0,")
            loConsulta.AppendLine("	C_Base = 0,")
            loConsulta.AppendLine("	C_Por_Impuesto = 0,")
            loConsulta.AppendLine("	C_Mon_Impuesto = 0,")
            loConsulta.AppendLine("            Ret_MontoRetenido = 0")
            loConsulta.AppendLine("WHERE Cod_Tip = 'FACT'")
            loConsulta.AppendLine("		AND status = 'Anulado'")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE #tmpTemporal")
            loConsulta.AppendLine("SET #tmpTemporal.Documento_Afect = dAfectado.doc_ori")
            loConsulta.AppendLine("FROM(	SELECT	renglones_dclientes.doc_ori,")
            loConsulta.AppendLine("				#tmpTemporal.documento")
            loConsulta.AppendLine("		FROM #tmpTemporal")
            loConsulta.AppendLine("			join devoluciones_clientes ON devoluciones_clientes.doc_des1 = #tmpTemporal.documento")
            loConsulta.AppendLine("										AND #tmpTemporal.Tipo_Documento = 'N/CR'")
            loConsulta.AppendLine("										AND #tmpTemporal.Documento_Afect = 'dev'")
            loConsulta.AppendLine("			JOIN renglones_dclientes ON renglones_dclientes.documento = devoluciones_clientes.documento")
            loConsulta.AppendLine("										AND renglones_dclientes.renglon = 1")
            loConsulta.AppendLine("	) AS dAfectado")
            loConsulta.AppendLine("WHERE #tmpTemporal.documento_afect = 'dev' and #tmpTemporal.documento = dAfectado.documento")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpTemporal(Fecha,Tipo,Documento,cod_tip,Reporte_Z,Tipo_Documento,Doc_Ini,Cod_Cli,Nom_Cli,Rif,Mon_Net,Mon_Exe,")
            loConsulta.AppendLine("							NC_Base,NC_Por_Impuesto,NC_Mon_Impuesto,C_Base,C_Por_Impuesto,C_Mon_Impuesto,")
            loConsulta.AppendLine("							Ret_Fecha,Ret_Comprobante,Ret_MontoRetenido,Doc_fin,Documento_Afect)")
            loConsulta.AppendLine("SELECT	Fecha,Tipo,'','','', Tipo_Documento,Doc_Ini,Cod_Cli,Nom_Cli,Rif,Mon_Net,Mon_Exe,")
            loConsulta.AppendLine("		NC_Base,NC_Por_Impuesto,NC_Mon_Impuesto,C_Base,C_Por_Impuesto,C_Mon_Impuesto,")
            loConsulta.AppendLine("		Ret_Fecha,Ret_Comprobante,Ret_MontoRetenido,'',doc")
            loConsulta.AppendLine("FROM")
            loConsulta.AppendLine("(SELECT		CAST(Cuentas_Cobrar.Fec_Ini AS DATE) 				AS Fecha,")
            loConsulta.AppendLine("			3													AS Tipo, ")
            loConsulta.AppendLine("                'C/RETENCIÓN'									    AS Tipo_Documento,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Num_Com							    AS Doc_Ini, ")
            loConsulta.AppendLine("			Clientes.Cod_Cli									AS Cod_Cli,")
            loConsulta.AppendLine("			Clientes.Nom_Cli    								AS Nom_Cli,")
            loConsulta.AppendLine("			Clientes.Rif										AS Rif,")
            loConsulta.AppendLine("			0								                    AS Mon_Net,")
            loConsulta.AppendLine("			0								                    AS Mon_Exe,")
            loConsulta.AppendLine("			0								                    AS NC_Base,")
            loConsulta.AppendLine("			0								                    AS NC_Por_Impuesto,")
            loConsulta.AppendLine("			0								                    AS NC_Mon_Impuesto,")
            loConsulta.AppendLine("			0								                    AS C_Base,")
            loConsulta.AppendLine("			0								                    AS C_Por_Impuesto,")
            loConsulta.AppendLine("			0								                    AS C_Mon_Impuesto,")
            loConsulta.AppendLine("			CAST(Cuentas_Cobrar.Fec_Ini AS DATE)                AS Ret_Fecha,")
            loConsulta.AppendLine("			/*CAST(YEAR(Cuentas_Cobrar.Fec_Ini)  AS VARCHAR(4))+")
            loConsulta.AppendLine("			RIGHT('00'+CAST(MONTH(Cuentas_Cobrar.Fec_Ini)  AS VARCHAR(4)),2)+")
            loConsulta.AppendLine("			RIGHT('00'+CAST(DAY(Cuentas_Cobrar.Fec_Ini)  AS VARCHAR(4)),2)+*/ retenciones_documentos.doc_ori						AS Ret_Comprobante,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Net                              AS Ret_MontoRetenido,")
            loConsulta.AppendLine("			Retenciones_Documentos.doc_ori						AS doc")
            loConsulta.AppendLine("FROM retenciones_documentos")
            loConsulta.AppendLine("	LEFT JOIN #tmpTemporal ON #tmpTemporal.documento = retenciones_documentos.doc_ori")
            loConsulta.AppendLine("	JOIN	cuentas_cobrar ")
            loConsulta.AppendLine("		ON	cuentas_cobrar.documento = retenciones_documentos.doc_des")
            loConsulta.AppendLine("		AND	cuentas_cobrar.cod_tip = retenciones_documentos.cla_des")
            loConsulta.AppendLine("	JOIN	Clientes ON Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loConsulta.AppendLine("WHERE	Cuentas_Cobrar.Status IN ('Pendiente','Afectado','Pagado')")
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Tip      =   'RETIVA' ")
            loConsulta.AppendLine("		AND retenciones_documentos.tip_ori = 'Cuentas_Cobrar' ")
            loConsulta.AppendLine("		AND	retenciones_documentos.Clase = 'IMPUESTO' ")
            loConsulta.AppendLine("		AND retenciones_documentos.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("		AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("		AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("		AND #tmpTemporal.Documento is null) AS Info ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("select * from #tmpTemporal ORDER BY	Fecha, Tipo_Documento, Doc_Ini,Documento_Afect")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DROP TABLE #tmpTemporal")

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")



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

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Ventas_GPV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLibro_Ventas_GPV.ReportSource = loObjetoReporte

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
' EAG: 22/09/15: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' EAG: 13/10/15: Se modificao informacion acerca de las retenciones de iva que no salian    '
'               como el comprobante de retencion, y retenciones hechas en meses anteriores. '
'-------------------------------------------------------------------------------------------'
