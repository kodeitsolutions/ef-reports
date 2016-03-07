'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Ventas_AT"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Ventas_AT
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
            'Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosFinales(6)
            'Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosFinales(7)
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()
            loConsulta.AppendLine("DECLARE @lnCero AS DECIMAL(28, 10);")
            loConsulta.AppendLine("SET @lnCero = CAST(0 AS DECIMAL(28, 10));")
            loConsulta.AppendLine("DECLARE @lcVacio AS NVARCHAR(30);")
            loConsulta.AppendLine("SET @lcVacio = N'';")
            loConsulta.AppendLine("")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpLibroVentas(	Operacion	        INT,")
            loConsulta.AppendLine("								Tabla	            VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Cod_Tip 	        VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Codigo_Tipo	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Documento	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Control	            VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Referencia	        VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Factura 	        VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Status		        VARCHAR(15) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Doc_Ori		        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Documento_Afectado	VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Cod_Cli		        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Fec_Ini		        DATETIME,")
            loConsulta.AppendLine("								Fec_Doc		        DATETIME,")
            loConsulta.AppendLine("								Tip_Doc		        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Mon_Bru 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Net 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Dis_Imp 	        XML,")
            loConsulta.AppendLine("								Mon_Des 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Rec 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Por_Des 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Por_Rec 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Otr1 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Otr2 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Otr3 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Com_Ret		    	VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Fec_Ret		    	DATETIME,")
            loConsulta.AppendLine("								Mon_Ret 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Imp 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Bas 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Por_Imp 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Exe 	        DECIMAL(28,10),")
            loConsulta.AppendLine("								Nom_Cli 	        VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Rif 	            VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Tip_Cli 	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Cli_Nacional 	    BIT,")
            loConsulta.AppendLine("								Periodo_Anterior 	BIT,")
            loConsulta.AppendLine("								Transaccion		    VARCHAR(10) COLLATE DATABASE_DEFAULT)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpLibroVentas(	Tabla, Cod_Tip, Codigo_Tipo, Documento, Control, Referencia, Factura, Status,")
            loConsulta.AppendLine("								Doc_Ori, Documento_Afectado, Cod_Cli, Fec_Ini, Fec_Doc, Tip_Doc, Mon_Bru, Mon_Net,")
            loConsulta.AppendLine("								Dis_Imp, Mon_Des, Mon_Rec, Por_Des, Por_Rec, Mon_Otr1, Mon_Otr2, Mon_Otr3,")
            loConsulta.AppendLine("								mon_ret, Mon_Imp, Mon_Bas, Por_Imp, Mon_Exe, ")
            loConsulta.AppendLine("								Nom_Cli, Rif, Tip_Cli, Cli_Nacional, Periodo_Anterior, Transaccion)")
            loConsulta.AppendLine("SELECT		")
            loConsulta.AppendLine("			'Ventas'															AS Tabla, 		")
            loConsulta.AppendLine("			CASE Cuentas_Cobrar.Cod_Tip															")
            loConsulta.AppendLine("			 	WHEN 'FACT' 	THEN 'Factura'													")
            loConsulta.AppendLine("			 	WHEN 'N/CR' 	THEN 'Nota de Credito'											")
            loConsulta.AppendLine("			 	WHEN 'N/DB' 	THEN 'Nota de Debito'											")
            loConsulta.AppendLine("			END																	AS Cod_Tip,		")
            loConsulta.AppendLine("			Cuentas_Cobrar.Cod_Tip												AS Codigo_Tipo,	")
            loConsulta.AppendLine("			CAST(Cuentas_Cobrar.Documento AS CHAR(30))							AS Documento, 	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Control												AS Control, 	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Referencia											AS Referencia,	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Factura												AS Factura, 	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Status												AS Status, 		")
            loConsulta.AppendLine("			Cuentas_Cobrar.Doc_Ori												AS Doc_Ori, 	")
            loConsulta.AppendLine("			@lcVacio															AS Documento_Afectado, 	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Cod_Cli												AS Cod_Cli, 	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Fec_Ini												AS Fec_Ini, 	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Fec_Doc												AS Fec_Doc, 	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Tip_Doc												AS Tip_Doc, 	")
            loConsulta.AppendLine("			CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Bru * -1 												")
            loConsulta.AppendLine("				ELSE Cuentas_Cobrar.Mon_Bru 														")
            loConsulta.AppendLine("			END																	AS Mon_Bru,  	")
            loConsulta.AppendLine("			CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Net * -1 												")
            loConsulta.AppendLine("				ELSE Cuentas_Cobrar.Mon_Net 														")
            loConsulta.AppendLine("			END																	AS Mon_Net,  	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Dis_Imp												AS Dis_Imp, 	")
            loConsulta.AppendLine("			CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Des * -1 												")
            loConsulta.AppendLine("				ELSE Cuentas_Cobrar.Mon_Des 														")
            loConsulta.AppendLine("			END																	AS Mon_Des,  	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Rec												AS Mon_Rec, 	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Por_Des												AS Por_Des, 	")
            loConsulta.AppendLine("			Cuentas_Cobrar.Por_Rec												AS Por_Rec, 	")
            loConsulta.AppendLine("			CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Otr1 * -1 												")
            loConsulta.AppendLine("				ELSE Cuentas_Cobrar.Mon_Otr1 													")
            loConsulta.AppendLine("			END																	AS Mon_Otr1,  	")
            loConsulta.AppendLine("			CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Otr2 * -1 												")
            loConsulta.AppendLine("				ELSE Cuentas_Cobrar.Mon_Otr2 													")
            loConsulta.AppendLine("			END																	AS Mon_Otr2,  	")
            loConsulta.AppendLine("			CASE WHEN Cuentas_Cobrar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("				THEN Cuentas_Cobrar.Mon_Otr3 * -1 												")
            loConsulta.AppendLine("				ELSE Cuentas_Cobrar.Mon_Otr3 													")
            loConsulta.AppendLine("			END																	AS Mon_Otr3,	")
            loConsulta.AppendLine("            @lnCero 															AS mon_ret, 	")
            loConsulta.AppendLine("            @lnCero 															AS mon_imp, 	")
            loConsulta.AppendLine("            @lnCero 															AS mon_bas, 	")
            loConsulta.AppendLine("            @lnCero 															AS por_imp, 	")
            loConsulta.AppendLine("            @lnCero 															AS mon_exe, ")
            loConsulta.AppendLine("			(CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') 				")
            loConsulta.AppendLine("				THEN Clientes.Nom_Cli 														")
            loConsulta.AppendLine("				ELSE (CASE WHEN (Cuentas_Cobrar.Nom_Cli = '') 									")
            loConsulta.AppendLine("					THEN Clientes.Nom_Cli 													")
            loConsulta.AppendLine("					ELSE Cuentas_Cobrar.Nom_Cli 													")
            loConsulta.AppendLine("				END) 																			")
            loConsulta.AppendLine("			END)																AS Nom_Cli, 	")
            loConsulta.AppendLine("			(CASE WHEN (Clientes.Generico = 0 AND Cuentas_Cobrar.Nom_Cli = '') 				")
            loConsulta.AppendLine("				THEN Clientes.Rif ELSE 														")
            loConsulta.AppendLine("			    (CASE WHEN (Cuentas_Cobrar.Rif = '')												")
            loConsulta.AppendLine("					THEN Clientes.Rif 														")
            loConsulta.AppendLine("					ELSE Cuentas_Cobrar.Rif 														")
            loConsulta.AppendLine("			    END) 																			")
            loConsulta.AppendLine("			 END)																AS Rif,			")
            loConsulta.AppendLine("			Clientes.Tip_Cli 												    AS Tip_Cli,		")
            loConsulta.AppendLine("			Clientes.Nacional 												    AS Cli_Nacional,		")
            loConsulta.AppendLine("			0																    AS Periodo_Anterior,		")
            loConsulta.AppendLine("			(CASE WHEN Cuentas_Cobrar.Status = 'Anulado' 										")
            loConsulta.AppendLine("				THEN '03-ANU' 																	")
            loConsulta.AppendLine("				ELSE '01-REG' 																	")
            loConsulta.AppendLine("			END)																AS Transaccion")
            loConsulta.AppendLine("FROM		Cuentas_Cobrar ")
            loConsulta.AppendLine("	JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli")
            loConsulta.AppendLine("WHERE		Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 			AND Cuentas_Cobrar.Documento BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro1Hasta)
            loConsulta.AppendLine(" 			AND Cuentas_Cobrar.Cod_Cli BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro2Hasta)
            loConsulta.AppendLine(" 			AND Cuentas_Cobrar.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro3Hasta)
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Status IN ( " & lcParametro6Desde & " ) ")
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
            loConsulta.AppendLine(" 			AND Cuentas_Cobrar.Cod_Tip IN ('FACT', 'N/CR', 'N/DB')")
            loConsulta.AppendLine("           AND Cuentas_Cobrar.Cod_Tip      BETWEEN " & lcParametro7Desde)
            loConsulta.AppendLine("           AND " & lcParametro7Hasta)
            loConsulta.AppendLine("")

            If CStr(lcParametro6Desde).ToLower().Contains("anulado") Then

                loConsulta.AppendLine("UNION ALL")

                loConsulta.AppendLine("")
                loConsulta.AppendLine("")

                '****************************************************************************************************
                '*************** Busca las Facturas de Venta Anuladas ******************************************
                '****************************************************************************************************

                loConsulta.AppendLine("")
                loConsulta.AppendLine("SELECT	")
                loConsulta.AppendLine("		'Facturas'													AS Tabla,		")
                loConsulta.AppendLine("		'Factura'													AS cod_tip,		")
                loConsulta.AppendLine("		'FACT'													    AS Codigo_Tipo,	")
                loConsulta.AppendLine("		CAST(Facturas.Documento AS CHAR(30))							AS Documento,	")
                loConsulta.AppendLine("		Facturas.Control												AS Control,		")
                loConsulta.AppendLine("		@lcVacio													AS Referencia,	")
                loConsulta.AppendLine("		Facturas.Factura												AS Factura,		")
                loConsulta.AppendLine("		Facturas.Status												AS Status,		")
                loConsulta.AppendLine("		@lcVacio													AS Doc_Ori, 	")
                loConsulta.AppendLine("		@lcVacio													AS Documento_Afectado, 	")
                loConsulta.AppendLine("		Facturas.Cod_Cli												AS Cod_Cli, 	")
                loConsulta.AppendLine("		Facturas.Fec_Ini												AS Fec_Ini, 	")
                loConsulta.AppendLine("		Facturas.Fec_Doc												AS Fec_Doc, 	")
                loConsulta.AppendLine("		@lcVacio													AS Tip_Doc, 	")
                loConsulta.AppendLine("		Facturas.Mon_Bru												AS Mon_Bru,		")
                loConsulta.AppendLine("		Facturas.Mon_Net												AS Mon_Net,		")
                loConsulta.AppendLine("		Facturas.Dis_imp												AS Dis_imp, 	")
                loConsulta.AppendLine("		Facturas.Mon_Des1											AS Mon_Des, 	")
                loConsulta.AppendLine("		Facturas.Mon_Rec1											AS Mon_Rec, 	")
                loConsulta.AppendLine("		Facturas.Por_Des1											AS Por_Des, 	")
                loConsulta.AppendLine("		Facturas.Por_Rec1											AS Por_Rec, 	")
                loConsulta.AppendLine("		Facturas.Mon_Otr1											AS Mon_Otr1, 	")
                loConsulta.AppendLine("		Facturas.Mon_Otr2											AS Mon_Otr2, 	")
                loConsulta.AppendLine("		Facturas.Mon_Otr3											AS Mon_Otr3,")
                loConsulta.AppendLine("		@lnCero 													AS mon_ret,		")
                loConsulta.AppendLine("		@lnCero 													AS mon_imp, 	")
                loConsulta.AppendLine("		@lnCero 													AS mon_bas, 	")
                loConsulta.AppendLine("		@lnCero 													AS por_imp, 	")
                loConsulta.AppendLine("		@lnCero 													AS mon_exe, ")
                loConsulta.AppendLine("		(CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '')				")
                loConsulta.AppendLine("			THEN Clientes.Nom_Cli												")
                loConsulta.AppendLine("			ELSE (CASE WHEN (Facturas.Nom_Cli = '')									")
                loConsulta.AppendLine("				THEN Clientes.Nom_Cli											")
                loConsulta.AppendLine("				ELSE Facturas.Nom_Cli												")
                loConsulta.AppendLine("			END)																	")
                loConsulta.AppendLine("		END)														AS  Nom_Cli,	")
                loConsulta.AppendLine("		(CASE WHEN (Clientes.Generico = 0 AND Facturas.Nom_Cli = '')				")
                loConsulta.AppendLine("			THEN Clientes.Rif													")
                loConsulta.AppendLine("			ELSE (CASE WHEN (Facturas.Rif = '')										")
                loConsulta.AppendLine("				THEN Clientes.Rif												")
                loConsulta.AppendLine("				ELSE Facturas.Rif													")
                loConsulta.AppendLine("			END)																	")
                loConsulta.AppendLine("		END)															AS  Rif,	")
                loConsulta.AppendLine("		Clientes.Tip_Cli												AS Tip_Cli,	")
                loConsulta.AppendLine("		Clientes.Nacional 											AS Cli_Nacional,		")
                loConsulta.AppendLine("		(CASE WHEN ( CONVERT(VARCHAR(6), Facturas.Fec_Ini, 112) > CONVERT(VARCHAR(6), Facturas.Fec_Doc, 112))")
                loConsulta.AppendLine("			THEN 1 																	")
                loConsulta.AppendLine("			ELSE 0 																	")
                loConsulta.AppendLine("		END)														    AS Periodo_Anterior,		")
                loConsulta.AppendLine("		(CASE WHEN Facturas.Status = 'Anulado'										")
                loConsulta.AppendLine("			THEN '03-ANU' 															")
                loConsulta.AppendLine("			ELSE '01-REG' 															")
                loConsulta.AppendLine("		END)														    AS Transaccion	")
                loConsulta.AppendLine("FROM		Facturas ")
                loConsulta.AppendLine("	JOIN	Clientes ON Facturas.Cod_Cli = Clientes.Cod_Cli ")
                loConsulta.AppendLine("WHERE		Facturas.Fec_Ini BETWEEN " & lcParametro0Desde)
                loConsulta.AppendLine(" 			AND " & lcParametro0Hasta)
                loConsulta.AppendLine(" 			AND Facturas.Documento BETWEEN " & lcParametro1Desde)
                loConsulta.AppendLine(" 			AND " & lcParametro1Hasta)
                loConsulta.AppendLine(" 			AND Facturas.Cod_Cli BETWEEN " & lcParametro2Desde)
                loConsulta.AppendLine(" 			AND " & lcParametro2Hasta)
                loConsulta.AppendLine(" 			AND Facturas.Cod_Suc BETWEEN " & lcParametro3Desde)
                loConsulta.AppendLine(" 			AND " & lcParametro3Hasta)
                loConsulta.AppendLine("			AND Facturas.Status  = 'Anulado'")
                If lcParametro5Desde = "Igual" Then
                    loConsulta.AppendLine(" 		AND Facturas.Cod_Rev BETWEEN " & lcParametro4Desde)
                Else
                    loConsulta.AppendLine(" 		AND Facturas.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
                End If
                loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
                loConsulta.AppendLine("")
                loConsulta.AppendLine("")

            End If


            '***************************************************************************
            ' Retenciones de IVA, si aplican 
            '***************************************************************************
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--Busca los datos de las retenciones")
            loConsulta.AppendLine("UPDATE		#tmpLibroVentas ")
            loConsulta.AppendLine("SET			Com_Ret = Retenciones.Com_Ret, ")
            loConsulta.AppendLine("			Fec_Ret = Retenciones.Fec_Ret, ")
            loConsulta.AppendLine("			Mon_Ret = Retenciones.Mon_Ret")
            loConsulta.AppendLine("FROM	(	SELECT		#tmpLibroVentas.Documento		AS Documento,")
            loConsulta.AppendLine("						#tmpLibroVentas.Cod_Tip		AS Cod_Tip,")
            loConsulta.AppendLine("						retenciones_documentos.num_com	AS Com_Ret,")
            loConsulta.AppendLine("						Cuentas_Cobrar.fec_ini			AS Fec_Ret,")
            loConsulta.AppendLine("						Cuentas_Cobrar.mon_net			AS Mon_Ret")
            loConsulta.AppendLine("			FROM		retenciones_documentos ")
            loConsulta.AppendLine("				JOIN	#tmpLibroVentas")
            loConsulta.AppendLine("					ON	#tmpLibroVentas.Documento = retenciones_documentos.doc_ori")
            loConsulta.AppendLine("					AND	#tmpLibroVentas.Codigo_Tipo = retenciones_documentos.cla_ori")
            loConsulta.AppendLine("					AND	retenciones_documentos.tip_ori = 'Cuentas_Cobrar'")
            loConsulta.AppendLine("					AND	retenciones_documentos.Clase = 'IMPUESTO'")
            loConsulta.AppendLine("				JOIN	Cuentas_Cobrar ")
            loConsulta.AppendLine("					ON	Cuentas_Cobrar.documento = retenciones_documentos.doc_des")
            loConsulta.AppendLine("					AND	Cuentas_Cobrar.cod_tip = retenciones_documentos.cla_des")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("		) AS Retenciones")
            loConsulta.AppendLine("WHERE	Retenciones.Documento = #tmpLibroVentas.documento")
            loConsulta.AppendLine("	AND	Retenciones.Cod_Tip = #tmpLibroVentas.Cod_Tip")
            loConsulta.AppendLine("")

            '***************************************************************************
            ' Datos de N/CR de devoluciones 
            '***************************************************************************
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE  #tmpLibroVentas")
            loConsulta.AppendLine("SET     Documento_Afectado = Facturas.Afectado")
            loConsulta.AppendLine("FROM    (   SELECT  RDev.Doc_Ori AS Afectado,")
            loConsulta.AppendLine("                    Dev.Doc_Des1 AS Factura ")
            loConsulta.AppendLine("            FROM    renglones_dClientes RDev")
            loConsulta.AppendLine("                JOIN devoluciones_Clientes Dev")
            loConsulta.AppendLine("                ON Dev.Documento = RDev.Documento")
            loConsulta.AppendLine("            WHERE   RDev.Tip_Ori = 'Facturas'")
            loConsulta.AppendLine("                AND Dev.Tip_Des1 = 'Cuentas_Cobrar'")
            loConsulta.AppendLine("                AND Dev.Cla_Des1 = 'N/CR'")
            loConsulta.AppendLine("        ) AS Facturas")
            loConsulta.AppendLine("WHERE   #tmpLibroVentas.Tabla = 'Ventas'")
            loConsulta.AppendLine("    AND #tmpLibroVentas.Codigo_Tipo = 'N/CR'")
            loConsulta.AppendLine("    AND #tmpLibroVentas.Documento = Facturas.Factura")
            loConsulta.AppendLine("")
            '***************************************************************************
            ' Genera el detalle de los impuestos 
            '***************************************************************************
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Genera el detalle de los impuestos")
            loConsulta.AppendLine("UPDATE  #tmpLibroVentas")
            loConsulta.AppendLine("SET     Mon_Imp = Impuestos.Mon_Imp*(CASE WHEN Impuestos.Codigo_Tipo='N/CR' THEN -1 ELSE 1 END),")
            loConsulta.AppendLine("        Por_Imp = Impuestos.Por_Imp,")
            loConsulta.AppendLine("        Mon_Exe = Impuestos.Mon_Exe*(CASE WHEN Impuestos.Codigo_Tipo='N/CR' THEN -1 ELSE 1 END),")
            loConsulta.AppendLine("        Mon_Bas = Impuestos.Mon_Bas*(CASE WHEN Impuestos.Codigo_Tipo='N/CR' THEN -1 ELSE 1 END)")
            loConsulta.AppendLine("FROM (  SELECT  Libro.Tabla, ")
            loConsulta.AppendLine("                Libro.Codigo_Tipo, ")
            loConsulta.AppendLine("                Libro.Documento, ")
            loConsulta.AppendLine("                MAX(CAST(T.C.value('porcentaje[1]', 'VARCHAR(MAX)') AS DECIMAL(28,10))) AS Por_Imp,")
            loConsulta.AppendLine("                SUM(CAST(T.C.value('base[1]', 'VARCHAR(MAX)') AS DECIMAL(28,10)))  AS Mon_Bas,")
            loConsulta.AppendLine("                SUM(CAST(T.C.value('exento[1]', 'VARCHAR(MAX)') AS DECIMAL(28,10)))  AS Mon_Exe,")
            loConsulta.AppendLine("                SUM(CAST(T.C.value('monto[1]', 'VARCHAR(MAX)') AS DECIMAL(28,10)))  AS Mon_Imp,")
            loConsulta.AppendLine("                MAX(Libro.Mon_Net) Mon_Net ")
            loConsulta.AppendLine("        FROM    #tmpLibroVentas AS Libro")
            loConsulta.AppendLine("	        CROSS APPLY Libro.Dis_Imp.nodes('//impuestos/impuesto') AS T(C)")
            loConsulta.AppendLine("        GROUP BY Libro.Tabla, Libro.Codigo_Tipo, Libro.Documento")
            loConsulta.AppendLine("        ) AS Impuestos")
            loConsulta.AppendLine("WHERE   #tmpLibroVentas.Tabla = Impuestos.Tabla")
            loConsulta.AppendLine("    AND #tmpLibroVentas.Codigo_Tipo = Impuestos.Codigo_Tipo")
            loConsulta.AppendLine("    AND #tmpLibroVentas.Documento = Impuestos.Documento")
            loConsulta.AppendLine("")

            '***************************************************************************
            ' Busca la fecha de los documentos del periodo anterior  
            '***************************************************************************
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Busca la fecha de los documentos del periodo anterior ")
            loConsulta.AppendLine("UPDATE  #tmpLibroVentas")
            loConsulta.AppendLine("SET     Fec_Doc = Facturas.Fec_Doc")   ',")
            'loConsulta.AppendLine("        Periodo_Anterior = (CASE WHEN ( CONVERT(VARCHAR(6), Facturas.Fec_Ini, 112) > CONVERT(VARCHAR(6), Facturas.Fec_Doc, 112)) THEN 1 ELSE 0 END)")
            loConsulta.AppendLine("FROM    Facturas")
            loConsulta.AppendLine("WHERE   Facturas.Documento = #tmpLibroVentas.Documento")
            loConsulta.AppendLine("    AND #tmpLibroVentas.Codigo_Tipo = 'FACT'")
            loConsulta.AppendLine("")

            '***************************************************************************
            ' Genera los números de operación 
            '***************************************************************************
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--Genera los números de operación")
            loConsulta.AppendLine("UPDATE  #tmpLibroVentas")
            loConsulta.AppendLine("SET     Operacion = P.Posicion")
            loConsulta.AppendLine("FROM    (   SELECT ROW_NUMBER() OVER (PARTITION BY Periodo_Anterior ORDER BY Fec_Ini) Posicion,")
            loConsulta.AppendLine("            Tabla, Codigo_Tipo, Documento")
            loConsulta.AppendLine("            FROM #tmpLibroVentas T")
            loConsulta.AppendLine("        ) P")
            loConsulta.AppendLine("WHERE   #tmpLibroVentas.Tabla = P.Tabla")
            loConsulta.AppendLine("    AND #tmpLibroVentas.Codigo_Tipo = P.Codigo_Tipo")
            loConsulta.AppendLine("    AND #tmpLibroVentas.Documento = P.Documento")
            loConsulta.AppendLine("")

            loConsulta.AppendLine("UPDATE #tmpLibroVentas")
            loConsulta.AppendLine("SET Mon_Net = 0,")
            loConsulta.AppendLine("	Mon_exe = 0,")
            loConsulta.AppendLine("	Mon_bas = 0,")
            loConsulta.AppendLine("	por_imp = 0,")
            loConsulta.AppendLine("	mon_imp = 0,")
            loConsulta.AppendLine(" mon_ret = 0")
            loConsulta.AppendLine("where #tmpLibroVentas.Transaccion = '03-ANU'")

            '***************************************************************************
            ' Consulta Final 
            '***************************************************************************
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  Tabla                                           AS Tabla,")
            loConsulta.AppendLine("        Status                                          AS Status,")
            loConsulta.AppendLine("        Codigo_Tipo                                     AS Codigo_Tipo,")
            loConsulta.AppendLine("        Operacion                                       AS Operacion,")
            loConsulta.AppendLine("        Periodo_Anterior                                AS Periodo_Anterior,")
            loConsulta.AppendLine("        COALESCE(Fec_Ret, Fec_Ini)                      AS Fec_Con,")
            loConsulta.AppendLine("        Fec_Ini                                         AS Fec_Ini,")
            loConsulta.AppendLine("        Fec_Doc                                         AS Fec_Doc,")
            loConsulta.AppendLine("        Rif                                             AS Rif,")
            loConsulta.AppendLine("        Nom_Cli                                         AS Nom_Cli,")
            loConsulta.AppendLine("        Cli_Nacional                                   AS Cli_Nacional,")
            loConsulta.AppendLine("        Com_Ret                                         AS Com_Ret,")
            loConsulta.AppendLine("        ''                                              AS Expediente_Importacion,")
            loConsulta.AppendLine("        (CASE WHEN Codigo_Tipo IN ('FACT') ")
            loConsulta.AppendLine("            THEN Documento ELSE '' END)                 AS Factura,")
            loConsulta.AppendLine("        Control,")
            loConsulta.AppendLine("        (CASE WHEN Codigo_Tipo IN ('N/DB') ")
            loConsulta.AppendLine("            THEN Documento ELSE '' END)                 AS Nota_Debito,")
            loConsulta.AppendLine("        (CASE WHEN Codigo_Tipo = 'N/CR' ")
            loConsulta.AppendLine("            THEN Documento ELSE '' END)                 AS Nota_Credito,")
            loConsulta.AppendLine("        Transaccion                                     AS Transaccion,")
            loConsulta.AppendLine("        Documento_Afectado                              AS Documento_Afectado,")
            loConsulta.AppendLine("        Mon_Net                                         AS Mon_Net,")
            loConsulta.AppendLine("        Mon_Exe                                         AS Mon_Exe,")
            loConsulta.AppendLine("        Mon_Bas                                         AS Mon_Bas,")
            loConsulta.AppendLine("        Por_Imp                                         AS Por_Imp,")
            loConsulta.AppendLine("        Mon_Imp                                         AS Mon_Imp,")
            loConsulta.AppendLine("        Mon_Ret                                         AS Mon_Ret,")
            loConsulta.AppendLine("        @lnCero                                         AS Mon_Ret_Terceros,")
            loConsulta.AppendLine("        @lnCero                                         AS Mon_Ret_Importacion")
            loConsulta.AppendLine("FROM    #tmpLibroVentas	")
            loConsulta.AppendLine("ORDER BY Periodo_Anterior ASC, Fec_Ini ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")




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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Ventas_AT", laDatosReporte)

            '-------------------------------------------------------------------
            ' Selección de opcion por excel (Microsoft Excel - xls):
            ' Genera el archivo a partir de la tabla de datos y termina la ejecución
            '-------------------------------------------------------------------
            If (Me.Request.QueryString("salida").ToLower = "xls") Then
                ' Ruta donde se creara temporalmente el archivo
                Dim lcFileName As String = Me.Server.MapPath("~\Administrativo\Temporales\rLibro_Ventas_AT_" & Guid.NewGuid().ToString("N") & ".xls")
                ' Se exporta para crear el archivo temporal
                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, lcFileName)

                ' Se modifica el contenido del archivo
                Me.mGenerarArchivoExcel(lcFileName, laDatosReporte.Tables(0), "")

                ' Se coloca en la respuesta para descargar
                Me.Response.Clear()
                'Me.Response.Buffer = True 
                Me.Response.AppendHeader("content-disposition", "attachment; filename=rLibro_Ventas_AT.xls")
                Me.Response.ContentType = "application/excel"
                Me.Response.WriteFile(lcFileName, True)
                'Me.Response.Write(Space(30))
                Me.Response.Flush()
                Me.Response.Close()

                Me.Response.End()

            End If


            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrrLibro_Ventas_AT.ReportSource = loObjetoReporte

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
            'Dim lcLogo As String = goEmpresa.pcUrlLogo 
            'lcLogo = HttpContext.Current.Server.MapPath(lcLogo)
            'loFormas = loHoja.Shapes

            'loFormas.AddPicture(lcLogo,  Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 1, 1, 60, 60)

            loRango = loHoja.Range("A1")
            loRango.Value = cusAplicacion.goEmpresa.pcNombre

            loRango = loHoja.Range("A2")
            loRango.Value = cusAplicacion.goEmpresa.pcRifEmpresa

            loRango = loHoja.Range("B5:AB5")
            loRango.Select()
            loRango.MergeCells = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.Value = "LIBRO DE VENTAS"
            loFuente = loRango.Font
            loFuente.Size = 14
            loFuente.Bold = True

            'Sub título del reporte
            Dim ldFechaReporte As Date
            loRango = loHoja.Range("B6:AB6")
            loRango.Select()
            loRango.MergeCells = True
            loRango.Value = "Mes de " & ldFechaReporte.ToString("MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-VE"))
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            ' Fecha y hora de creacion
            Dim ldFecha As DateTime = Date.Now()
            loRango = loHoja.Range("AB1")
            loRango.NumberFormat = "@"
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            loRango.Value = ldFecha.ToString("dd/MM/yyyy")

            loRango = loHoja.Range("AB2")
            loRango.NumberFormat = "@" 'La celda almacena un string
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            loRango.Value = ldFecha.ToString("hh:mm:ss tt")

            ' Parametros del reporte
            'loRango = loHoja.Range("B7:O7")
            'loRango.Select()
            'loRango.MergeCells = True
            'loRango.Value = lcParametrosReporte
            'loRango.WrapText = True
            'loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify


            Dim lnFilaActual As Integer = 8

            '******************************************************************'
            ' Datos del Reporte												'
            '******************************************************************'
            loRango = loHoja.Range("B" & lnFilaActual)
            loRango.Value = "Oper." & vbLf & "Nro."

            loRango = loHoja.Range("C" & lnFilaActual)
            loRango.Value = "Fecha" & vbLf & "Contab."

            loRango = loHoja.Range("D" & lnFilaActual)
            loRango.Value = "Fecha de" & vbLf & "la Factura"

            loRango = loHoja.Range("E" & lnFilaActual)
            loRango.Value = "RIF"

            loRango = loHoja.Range("F" & lnFilaActual)
            loRango.Value = "Nombre o Razón Social"

            loRango = loHoja.Range("G" & lnFilaActual)
            loRango.Value = "Núm. de" & vbLf & "Planilla" & vbLf & "Importación"

            loRango = loHoja.Range("H" & lnFilaActual)
            loRango.Value = "Valor FOB" & vbLf & "de la " & vbLf & "Mercancia"

            loRango = loHoja.Range("I" & lnFilaActual)
            loRango.Value = "Número de" & vbLf & "Factura"

            loRango = loHoja.Range("J" & lnFilaActual)
            loRango.Value = "Número de" & vbLf & "Control de" & vbLf & "Factura"

            loRango = loHoja.Range("K" & lnFilaActual)
            loRango.Value = "Número de" & vbLf & "Nota de" & vbLf & "Débito"

            loRango = loHoja.Range("L" & lnFilaActual)
            loRango.Value = "Número de" & vbLf & "Nota de" & vbLf & "Crédito"

            loRango = loHoja.Range("M" & lnFilaActual)
            loRango.Value = "Tipo de" & vbLf & "Transac."

            loRango = loHoja.Range("N" & lnFilaActual)
            loRango.Value = "Número de" & vbLf & "Factura" & vbLf & "Afectada"

            loRango = loHoja.Range("O" & lnFilaActual)
            loRango.Value = "Total" & vbLf & "Ventas" & vbLf & "Incl. IVA"

            loRango = loHoja.Range("P" & lnFilaActual)
            loRango.Value = "Ventas" & vbLf & "Exentas"

            loRango = loHoja.Range("Q" & lnFilaActual)
            loRango.Value = "Ventas" & vbLf & "Exoneradas"

            loRango = loHoja.Range("R" & lnFilaActual)
            loRango.Value = "Ventas" & vbLf & "No Sujetas"

            loRango = loHoja.Range("S" & lnFilaActual)
            loRango.Value = "Ventas Internas" & vbLf & "No Gravadas"

            loRango = loHoja.Range("T" & lnFilaActual)
            loRango.Value = "Base" & vbLf & "Imponible"

            loRango = loHoja.Range("U" & lnFilaActual)
            loRango.Value = "%" & vbLf & "Alic."

            loRango = loHoja.Range("V" & lnFilaActual)
            loRango.Value = "Impuesto" & vbLf & "IVA"

            loRango = loHoja.Range("W" & lnFilaActual)
            loRango.Value = "IVA Retenido" & vbLf & "(por el comprador)"

            loRango = loHoja.Range("X" & lnFilaActual)
            loRango.Value = "Base" & vbLf & "Imponible"

            loRango = loHoja.Range("Y" & lnFilaActual)
            loRango.Value = "%" & vbLf & "Alic."

            loRango = loHoja.Range("Z" & lnFilaActual)
            loRango.Value = "Impuesto" & vbLf & "IVA"

            loRango = loHoja.Range("AA" & lnFilaActual)
            loRango.Value = "IVA Percibido"

            loRango = loHoja.Range("B" & lnFilaActual & ":AB" & lnFilaActual)
            loFuente = loRango.Font
            loFuente.Bold = True
            'loFuente.Color = Rgb(255, 255, 255)
            loRango.Interior.Color = Rgb(200, 200, 200)

            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            '****************************************************************************************
            ' Facturas del Periodo actual
            '****************************************************************************************

            Dim lnFilaInicio As Integer = lnFilaActual
            Dim laRenglones() As DataRow = loDatos.Select("Periodo_Anterior=0")
            For Each loRenglon As DataRow In laRenglones
                lnFilaActual += 1

                'Oper. Nro."
                loRango = loHoja.Range("B" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CInt(loRenglon("Operacion"))
                loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

                'Fecha Contab.
                loRango = loHoja.Range("C" & lnFilaActual)
                loRango.NumberFormat = "dd-mm-yyyy;@"
                loRango.Value = CDate(loRenglon("Fec_Con"))
                loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                'Fecha de la Factura
                loRango = loHoja.Range("D" & lnFilaActual)
                loRango.NumberFormat = "dd-mm-yyyy;@"
                loRango.Value = CDate(loRenglon("Fec_Ini"))
                loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

                'RIF
                loRango = loHoja.Range("E" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Rif")).Trim()

                'Nombre o Razón Social 
                loRango = loHoja.Range("F" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Nom_Cli")).Trim()

                'Número Comprobante
                loRango = loHoja.Range("G" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Expediente_Importacion")).Trim()

                'Núm. de Expediente Importación
                loRango = loHoja.Range("H" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = CDec(loRenglon("Mon_Net"))

                'Número de Factura
                loRango = loHoja.Range("I" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Factura")).Trim()

                'Número de Control de Factura
                loRango = loHoja.Range("J" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Control")).Trim()

                'Número de Nota de Débito
                loRango = loHoja.Range("K" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Nota_Debito")).Trim()

                'Número de Nota de Crédito
                loRango = loHoja.Range("L" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Nota_Credito")).Trim()

                'Tipo de Transac.
                loRango = loHoja.Range("M" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Transaccion")).Trim()

                'Número de Factura Afectada
                loRango = loHoja.Range("N" & lnFilaActual)
                loRango.NumberFormat = "@"
                loRango.Value = CStr(loRenglon("Documento_Afectado")).Trim()

                'Total Ventas Incl. IVA
                loRango = loHoja.Range("O" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = CDec(loRenglon("Mon_Net"))

                'Ventas Exentas
                loRango = loHoja.Range("P" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = CDec(loRenglon("Mon_Exe"))

                'Ventas Exoneradas
                loRango = loHoja.Range("Q" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = CDec(0)

                'Ventas No Sujetas
                loRango = loHoja.Range("R" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = CDec(0)

                'Ventas Internas No Gravadas
                loRango = loHoja.Range("S" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = CDec(loRenglon("Mon_Exe"))

                'Base Imponible
                Dim llEsNacional As Boolean = CBool(loRenglon("Cli_Nacional"))
                loRango = loHoja.Range("T" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = IIf(llEsNacional, CDec(loRenglon("Mon_Bas")), CDec(0))

                '% Alic.
                Dim lnPorcentajeImpuesto As Decimal = CDec(loRenglon("Por_Imp"))
                loRango = loHoja.Range("U" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = IIf(llEsNacional, lnPorcentajeImpuesto, CDec(0))

                'Impuesto IVA
                loRango = loHoja.Range("V" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = IIf(llEsNacional, CDec(loRenglon("Mon_Imp")), CDec(0))

                'IVA Retenido (por el comprador)
                loRango = loHoja.Range("W" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = IIf(llEsNacional, CDec(loRenglon("Mon_Ret")), CDec(0))

                'Base Imponible
                loRango = loHoja.Range("X" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = IIf(Not llEsNacional, CDec(loRenglon("Mon_Bas")), CDec(0))

                '% Alic.
                loRango = loHoja.Range("Y" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = IIf(Not llEsNacional, lnPorcentajeImpuesto, CDec(0))

                'Impuesto IVA
                loRango = loHoja.Range("Z" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = IIf(Not llEsNacional, CDec(loRenglon("Mon_Imp")), CDec(0))

                'IVA Retenido (por el comprador)
                loRango = loHoja.Range("AA" & lnFilaActual)
                loRango.NumberFormat = lcFormatoMontos
                loRango.Value = IIf(Not llEsNacional, CDec(loRenglon("Mon_Ret")), CDec(0))

                'Alicuota
                loRango = loHoja.Range("AB" & lnFilaActual)
                loRango.NumberFormat = "@"
                If (CStr(loRenglon("Status")).ToLower().Trim() = "anulado") Then
                    loRango.Value = "ANULADO"
                Else
                    Dim lcTipoAlicuota As String
                    lcTipoAlicuota = IIf(CBool(loRenglon("Cli_Nacional")), "INTERNA", "EXPORTACION")
                    If (lnPorcentajeImpuesto = 0D) Then
                        lcTipoAlicuota = lcTipoAlicuota & "-EXENTO"
                    ElseIf lnPorcentajeImpuesto < 12D Then
                        lcTipoAlicuota = lcTipoAlicuota & "-REDUCIDA"
                    ElseIf lnPorcentajeImpuesto = 12D Then
                        lcTipoAlicuota = lcTipoAlicuota & "-GENERAL"
                    Else 'If lnPorcentajeImpuesto > 12D 
                        lcTipoAlicuota = lcTipoAlicuota & "-ADICIONAL"
                    End If
                    loRango.Value = lcTipoAlicuota
                End If

            Next loRenglon

            Dim lnTotal As Integer = laRenglones.Length
            loRango = loHoja.Range("B" & (lnFilaInicio) & ":AA" & (lnFilaInicio))
            loRango.Select()
            loExcel.Selection.AutoFilter()

            loRango = loHoja.Range("B" & (lnFilaInicio + 1) & ":AA" & (lnFilaInicio + lnTotal))
            loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            Dim lnDesde As Integer = lnFilaInicio + 1
            Dim lnHasta As Integer = lnFilaInicio + lnTotal

            lnFilaInicio += lnTotal + 2
            loRango = loHoja.Range("B" & (lnFilaInicio) & ":C" & (lnFilaInicio))
            loRango.MergeCells = True
            loRango.NumberFormat = "@"
            loRango.Value = "Total Registros: " & lnTotal.ToString()
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            loRango = loHoja.Range("G" & (lnFilaInicio))
            loRango.NumberFormat = "@"
            loRango.Value = "Total General: "
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

            loRango = loHoja.Range("H" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", H" & lnDesde & ":H" & lnHasta & ")"

            loRango = loHoja.Range("O" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", O" & lnDesde & ":O" & lnHasta & ")"

            loRango = loHoja.Range("P" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", P" & lnDesde & ":P" & lnHasta & ")"

            loRango = loHoja.Range("Q" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", Q" & lnDesde & ":Q" & lnHasta & ")"

            loRango = loHoja.Range("R" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", R" & lnDesde & ":R" & lnHasta & ")"

            loRango = loHoja.Range("S" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", S" & lnDesde & ":S" & lnHasta & ")"

            loRango = loHoja.Range("T" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", T" & lnDesde & ":T" & lnHasta & ")"

            loRango = loHoja.Range("V" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", V" & lnDesde & ":V" & lnHasta & ")"

            loRango = loHoja.Range("W" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", W" & lnDesde & ":W" & lnHasta & ")"

            loRango = loHoja.Range("X" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", X" & lnDesde & ":X" & lnHasta & ")"

            loRango = loHoja.Range("Z" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", Z" & lnDesde & ":Z" & lnHasta & ")"

            loRango = loHoja.Range("AA" & (lnFilaInicio))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", AA" & lnDesde & ":AA" & lnHasta & ")"

            loRango = loHoja.Range("B" & (lnFilaInicio) & ":AB" & (lnFilaInicio))
            loFuente = loRango.Font
            loFuente.Bold = True

            '****************************************************************************************
            ' Bloque de totales
            '****************************************************************************************
            lnFilaActual = lnFilaActual + 4
            loRango = loHoja.Range("K" & (lnFilaActual))
            loRango.NumberFormat = "@"
            loRango.Value = "Base Imponible"
            loFuente = loRango.Font
            loFuente.Bold = True

            loRango = loHoja.Range("L" & (lnFilaActual))
            loRango.NumberFormat = "@"
            loRango.Value = "Credito Fiscal"
            loFuente = loRango.Font
            loFuente.Bold = True

            loRango = loHoja.Range("M" & (lnFilaActual))
            loRango.NumberFormat = "@"
            loRango.Value = "IVA Retenido"
            loFuente = loRango.Font
            loFuente.Bold = True

            lnFilaActual = lnFilaActual + 1
            loRango = loHoja.Range("G" & (lnFilaActual))
            loRango.NumberFormat = "@"
            loRango.Value = "Total Ventas Internas No Gravadas"
            loFuente = loRango.Font
            loFuente.Bold = True

            loRango = loHoja.Range("K" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""<>ANULADO"", S" & lnDesde & ":S" & lnHasta & ")"

            loRango = loHoja.Range("L" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = 0 '"=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", P" & lnDesde & ":P" & lnHasta	& ")"

            loRango = loHoja.Range("M" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = 0 '"=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", P" & lnDesde & ":P" & lnHasta	& ")"


            lnFilaActual = lnFilaActual + 1
            loRango = loHoja.Range("G" & (lnFilaActual))
            loRango.NumberFormat = "@"
            loRango.Value = "Total Ventas de Exportación"
            loFuente = loRango.Font
            loFuente.Bold = True

            loRango = loHoja.Range("K" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=EXPORTACION-GENERAL"", X" & lnDesde & ":X" & lnHasta & ") + " & _
                              " SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=EXPORTACION-ADICIONAL"", X" & lnDesde & ":X" & lnHasta & ") + " & _
                              " SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=EXPORTACION-REDUCIDA"", X" & lnDesde & ":X" & lnHasta & ") "

            loRango = loHoja.Range("L" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=EXPORTACION-GENERAL"", Y" & lnDesde & ":Y" & lnHasta & ") + " & _
                              " SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=EXPORTACION-ADICIONAL"", Y" & lnDesde & ":Y" & lnHasta & ") + " & _
                              " SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=EXPORTACION-REDUCIDA"", Y" & lnDesde & ":Y" & lnHasta & ") "

            loRango = loHoja.Range("M" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=EXPORTACION-GENERAL"", Z" & lnDesde & ":Z" & lnHasta & ") + " & _
                              " SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=EXPORTACION-ADICIONAL"", Z" & lnDesde & ":Z" & lnHasta & ") + " & _
                              " SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=EXPORTACION-REDUCIDA"", Z" & lnDesde & ":Z" & lnHasta & ") "


            lnFilaActual = lnFilaActual + 1
            loRango = loHoja.Range("G" & (lnFilaActual))
            loRango.NumberFormat = "@"
            loRango.Value = "Total Ventas Internas Afectas solo Alicuota General"
            loFuente = loRango.Font
            loFuente.Bold = True

            loRango = loHoja.Range("K" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=INTERNA-GENERAL"", T" & lnDesde & ":T" & lnHasta & ")"

            loRango = loHoja.Range("L" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=INTERNA-GENERAL"", V" & lnDesde & ":V" & lnHasta & ")"

            loRango = loHoja.Range("M" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=INTERNA-GENERAL"", W" & lnDesde & ":W" & lnHasta & ")"

            lnFilaActual = lnFilaActual + 1
            loRango = loHoja.Range("G" & (lnFilaActual))
            loRango.NumberFormat = "@"
            loRango.Value = "Total Compras Internas Afectas en Alícuota General + Adicional"
            loFuente = loRango.Font
            loFuente.Bold = True

            loRango = loHoja.Range("K" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=INTERNA-ADICIONAL"", T" & lnDesde & ":T" & lnHasta & ")"

            loRango = loHoja.Range("L" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=INTERNA-ADICIONAL"", V" & lnDesde & ":V" & lnHasta & ")"

            loRango = loHoja.Range("M" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=INTERNA-ADICIONAL"", W" & lnDesde & ":W" & lnHasta & ")"

            lnFilaActual = lnFilaActual + 1
            loRango = loHoja.Range("G" & (lnFilaActual))
            loRango.NumberFormat = "@"
            loRango.Value = "Total Compras Internas Afectas en Alícuota Reducida"
            loFuente = loRango.Font
            loFuente.Bold = True

            loRango = loHoja.Range("K" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=INTERNA-REDUCIDA"", T" & lnDesde & ":T" & lnHasta & ")"

            loRango = loHoja.Range("L" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=INTERNA-REDUCIDA"", V" & lnDesde & ":V" & lnHasta & ")"

            loRango = loHoja.Range("M" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta & ", ""=INTERNA-REDUCIDA"", W" & lnDesde & ":W" & lnHasta & ")"

            lnFilaActual = lnFilaActual + 1
            lnDesde = lnFilaActual - 5
            lnHasta = lnFilaActual - 1

            loRango = loHoja.Range("K" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUM(K" & lnDesde & ":K" & lnHasta & ")"
            loFuente = loRango.Font
            loFuente.Bold = True

            loRango = loHoja.Range("L" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUM(L" & lnDesde & ":L" & lnHasta & ")"
            loFuente = loRango.Font
            loFuente.Bold = True

            loRango = loHoja.Range("M" & (lnFilaActual))
            loRango.NumberFormat = lcFormatoMontos
            loRango.Formula = "=SUM(M" & lnDesde & ":M" & lnHasta & ")"
            loFuente = loRango.Font
            loFuente.Bold = True


            '     '****************************************************************************************
            '     ' Facturas del Periodo anterior
            '     '****************************************************************************************
            'lnFilaActual = lnFilaActual + 3

            'loRango = loHoja.Range("B" & lnFilaActual)
            'loFuente = loRango.Font
            'loFuente.Bold = True
            'loFuente.Size = 14
            'loRango.Value = "AJUSTES"

            'lnFilaActual = lnFilaActual + 1

            'loRango = loHoja.Range("B" & lnFilaActual)
            'loRango.Value = "Oper." & vbLf & "Nro."

            'loRango = loHoja.Range("C" & lnFilaActual)
            'loRango.Value = "Fecha" & vbLf & "Contab."

            'loRango = loHoja.Range("D" & lnFilaActual)
            'loRango.Value = "Fecha de" & vbLf & "la Factura"

            'loRango = loHoja.Range("E" & lnFilaActual)
            'loRango.Value = "RIF"

            'loRango = loHoja.Range("F" & lnFilaActual)
            'loRango.Value = "Nombre o Razón Social"

            'loRango = loHoja.Range("G" & lnFilaActual)
            'loRango.Value = "Núm. de" & vbLf & "Planilla" & vbLf & "Importación"

            'loRango = loHoja.Range("H" & lnFilaActual)
            'loRango.Value = "Valor FOV" & vbLf & "de la " & vbLf & "Mercancia"

            'loRango = loHoja.Range("I" & lnFilaActual)
            'loRango.Value = "Número de" & vbLf & "Factura"

            'loRango = loHoja.Range("J" & lnFilaActual)
            'loRango.Value = "Número de" & vbLf & "Control de" & vbLf & "Factura"

            '         loRango = loHoja.Range("K" & lnFilaActual)
            'loRango.Value = "Número de" & vbLf & "Nota de" & vbLf & "Débito"

            '         loRango = loHoja.Range("L" & lnFilaActual)
            'loRango.Value = "Número de" & vbLf & "Nota de" & vbLf & "Crédito"

            '         loRango = loHoja.Range("M" & lnFilaActual)
            'loRango.Value = "Tipo de" & vbLf & "Transac."

            '         loRango = loHoja.Range("N" & lnFilaActual)
            'loRango.Value = "Número de" & vbLf & "Factura" & vbLf & "Afectada"

            '         loRango = loHoja.Range("O" & lnFilaActual)
            'loRango.Value = "Total" & vbLf & "Ventas" & vbLf & "Incl. IVA"

            '         loRango = loHoja.Range("P" & lnFilaActual)
            'loRango.Value = "Ventas" & vbLf & "Exentas"

            '         loRango = loHoja.Range("Q" & lnFilaActual)
            'loRango.Value = "Ventas" & vbLf & "Exoneradas"

            '         loRango = loHoja.Range("R" & lnFilaActual)
            'loRango.Value = "Ventas" & vbLf & "No Sujetas"

            '         loRango = loHoja.Range("S" & lnFilaActual)
            'loRango.Value = "Ventas Internas" & vbLf & "No Gravadas"

            '         loRango = loHoja.Range("T" & lnFilaActual)
            'loRango.Value = "Base" & vbLf & "Imponible"

            '         loRango = loHoja.Range("U" & lnFilaActual)
            'loRango.Value = "%" & vbLf & "Alic."

            '         loRango = loHoja.Range("V" & lnFilaActual)
            'loRango.Value = "Impuesto" & vbLf & "IVA"

            '         loRango = loHoja.Range("W" & lnFilaActual)
            'loRango.Value = "IVA Retenido" & vbLf & "(por el comprador)" 

            '         loRango = loHoja.Range("X" & lnFilaActual)
            'loRango.Value = "Base" & vbLf & "Imponible"

            '         loRango = loHoja.Range("Y" & lnFilaActual)
            'loRango.Value = "%" & vbLf & "Alic."

            '         loRango = loHoja.Range("Z" & lnFilaActual)
            'loRango.Value = "Impuesto" & vbLf & "IVA"

            '         loRango = loHoja.Range("AA" & lnFilaActual)
            'loRango.Value = "IVA Percibido" 

            '         loRango = loHoja.Range("B" & lnFilaActual & ":AB" & lnFilaActual)
            'loFuente = loRango.Font
            'loFuente.Bold = True
            ''loFuente.Color = Rgb(255, 255, 255)
            'loRango.Interior.Color = Rgb(200, 200, 200)

            'loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            'loRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            'loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


            'lnFilaInicio = lnFilaActual
            '         laRenglones = loDatos.Select("Periodo_Anterior=1")
            '         For Each loRenglon As DataRow In laRenglones

            '          lnFilaActual += 1

            '          'Oper. Nro."
            '          loRango = loHoja.Range("B" & lnFilaActual)
            '          loRango.NumberFormat = "@"
            '          loRango.Value = CInt(loRenglon("Operacion"))
            '          loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

            '          'Fecha Contab.
            '          loRango = loHoja.Range("C" & lnFilaActual)
            '          loRango.NumberFormat = "dd-mm-yyyy;@"
            '          loRango.Value = CDate(loRenglon("Fec_Con"))
            '          loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            '          'Fecha de la Factura
            '          loRango = loHoja.Range("D" & lnFilaActual)
            '          loRango.NumberFormat = "dd-mm-yyyy;@"
            '          loRango.Value = CDate(loRenglon("Fec_Ini"))
            '          loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

            '          'RIF
            '          loRango = loHoja.Range("E" & lnFilaActual)
            '          loRango.NumberFormat = "@"
            '          loRango.Value = CStr(loRenglon("Rif")).Trim()

            '          'Nombre o Razón Social 
            '          loRango = loHoja.Range("F" & lnFilaActual)	
            '          loRango.NumberFormat = "@"
            '          loRango.Value = CStr(loRenglon("Nom_Cli")).Trim()

            '          'Número Comprobante
            '          loRango = loHoja.Range("G" & lnFilaActual) 
            '          loRango.NumberFormat = "@"
            '          loRango.Value = CStr(loRenglon("Expediente_Importacion")).Trim()

            '          'Núm. de Expediente Importación
            '          loRango = loHoja.Range("H" & lnFilaActual) 
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = CDec(loRenglon("Mon_Net"))

            '          'Número de Factura
            '          loRango = loHoja.Range("I" & lnFilaActual) 
            '          loRango.NumberFormat = "@"
            '          loRango.Value = CStr(loRenglon("Factura")).Trim()

            '          'Número de Control de Factura
            '          loRango = loHoja.Range("J" & lnFilaActual)   
            '          loRango.NumberFormat = "@"
            '          loRango.Value = CStr(loRenglon("Control")).Trim()

            '             'Número de Nota de Débito
            '             loRango = loHoja.Range("K" & lnFilaActual)
            '          loRango.NumberFormat = "@"
            '          loRango.Value = CStr(loRenglon("Nota_Debito")).Trim()

            '             'Número de Nota de Crédito
            '             loRango = loHoja.Range("L" & lnFilaActual)
            '             loRango.NumberFormat = "@"
            '             loRango.Value = CStr(loRenglon("Nota_Credito")).Trim()

            '          'Tipo de Transac.
            '             loRango = loHoja.Range("M" & lnFilaActual)
            '          loRango.NumberFormat = "@"
            '             loRango.Value = CStr(loRenglon("Transaccion")).Trim()

            '          'Número de Factura Afectada
            '             loRango = loHoja.Range("N" & lnFilaActual)
            '             loRango.NumberFormat = "@"
            '             loRango.Value = CStr(loRenglon("Documento_Afectado")).Trim()

            '          'Total Ventas Incl. IVA
            '             loRango = loHoja.Range("O" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = CDec(loRenglon("Mon_Net"))

            '          'Ventas Exentas
            '             loRango = loHoja.Range("P" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = CDec(loRenglon("Mon_Exe"))

            '          'Ventas Exoneradas
            '             loRango = loHoja.Range("Q" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = CDec(0)

            '          'Ventas No Sujetas
            '             loRango = loHoja.Range("R" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = CDec(0)

            '          'Ventas Internas No Gravadas
            '             loRango = loHoja.Range("S" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = CDec(loRenglon("Mon_Exe"))

            '          'Base Imponible
            '             Dim llEsNacional As Boolean = CBool(loRenglon("Cli_Nacional"))
            '             loRango = loHoja.Range("T" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = IIf(llEsNacional, CDec(loRenglon("Mon_Bas")), CDec(0))

            '          '% Alic.
            '             Dim lnPorcentajeImpuesto As Decimal = CDec(loRenglon("Por_Imp"))
            '             loRango = loHoja.Range("U" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = IIf(llEsNacional, lnPorcentajeImpuesto, CDec(0)) 

            '          'Impuesto IVA
            '             loRango = loHoja.Range("V" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = IIf(llEsNacional, CDec(loRenglon("Mon_Imp")), CDec(0)) 

            '          'IVA Retenido (por el comprador)
            '             loRango = loHoja.Range("W" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = IIf(llEsNacional, CDec(loRenglon("Mon_Ret")), CDec(0)) 

            '          'Base Imponible
            '             loRango = loHoja.Range("X" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = IIf(Not llEsNacional, CDec(loRenglon("Mon_Bas")), CDec(0))

            '          '% Alic.
            '             loRango = loHoja.Range("Y" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = IIf(Not llEsNacional, lnPorcentajeImpuesto, CDec(0)) 

            '          'Impuesto IVA
            '             loRango = loHoja.Range("Z" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = IIf(Not llEsNacional, CDec(loRenglon("Mon_Imp")), CDec(0)) 

            '          'IVA Retenido (por el comprador)
            '             loRango = loHoja.Range("AA" & lnFilaActual)
            '          loRango.NumberFormat = lcFormatoMontos
            '             loRango.Value = IIf(Not llEsNacional, CDec(loRenglon("Mon_Ret")), CDec(0)) 

            '          'Alicuota
            '             loRango = loHoja.Range("AB" & lnFilaActual)
            '          loRango.NumberFormat = "@"
            '             If (CStr(loRenglon("Status")).ToLower().Trim() = "anulado") Then
            '                 loRango.Value = "ANULADO"
            '             Else 
            '                 Dim lcTipoAlicuota As String 
            '                 lcTipoAlicuota = IIf(cbool(loRenglon("Cli_Nacional")), "INTERNA", "EXPORTACION")
            '                 If (lnPorcentajeImpuesto = 0D) Then
            '                     lcTipoAlicuota = lcTipoAlicuota & "-EXENTO"
            '                 ElseIf lnPorcentajeImpuesto < 12D 
            '                     lcTipoAlicuota = lcTipoAlicuota & "-REDUCIDA"
            '                 ElseIf lnPorcentajeImpuesto = 12D 
            '                     lcTipoAlicuota = lcTipoAlicuota & "-GENERAL"
            '                 Else 'If lnPorcentajeImpuesto > 12D 
            '                     lcTipoAlicuota = lcTipoAlicuota & "-ADICIONAL"
            '                 End If
            '                 loRango.Value = lcTipoAlicuota
            '             End If

            '         Next loRenglon


            'lnTotal = laRenglones.Length

            '         loRango = loHoja.Range("B" & (lnFilaInicio + 1) & ":AA" & (lnFilaInicio + lnTotal))
            'loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

            'lnDesde = lnFilaInicio
            'lnHasta = lnFilaInicio + lnTotal

            'lnFilaInicio += lnTotal + 2
            'loRango = loHoja.Range("B" & (lnFilaInicio) & ":C" & (lnFilaInicio))
            'loRango.MergeCells = True
            'loRango.NumberFormat = "@"
            'loRango.Value = "Total Registros: " & lnTotal.ToString()
            'loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

            'loRango = loHoja.Range("G" & (lnFilaInicio))
            'loRango.NumberFormat = "@"
            'loRango.Value = "Total General: "
            'loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

            'loRango = loHoja.Range("H" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", H" & lnDesde & ":H" & lnHasta	& ")"

            'loRango = loHoja.Range("O" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", O" & lnDesde & ":O" & lnHasta	& ")"

            'loRango = loHoja.Range("P" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", P" & lnDesde & ":P" & lnHasta	& ")"

            'loRango = loHoja.Range("Q" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", Q" & lnDesde & ":Q" & lnHasta	& ")"

            'loRango = loHoja.Range("R" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", R" & lnDesde & ":R" & lnHasta	& ")"

            'loRango = loHoja.Range("S" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", S" & lnDesde & ":S" & lnHasta	& ")"

            'loRango = loHoja.Range("T" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", T" & lnDesde & ":T" & lnHasta	& ")"

            'loRango = loHoja.Range("V" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", V" & lnDesde & ":V" & lnHasta	& ")"

            'loRango = loHoja.Range("W" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", W" & lnDesde & ":W" & lnHasta	& ")"

            'loRango = loHoja.Range("X" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", X" & lnDesde & ":X" & lnHasta	& ")"

            'loRango = loHoja.Range("Z" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", Z" & lnDesde & ":Z" & lnHasta	& ")"

            'loRango = loHoja.Range("AA" & (lnFilaInicio))
            'loRango.NumberFormat = lcFormatoMontos
            '         loRango.Formula = "=SUMIF(AB" & lnDesde & ":AB" & lnHasta	& ", ""<>ANULADO"", AA" & lnDesde & ":AA" & lnHasta	& ")"

            'loRango = loHoja.Range("B" & (lnFilaInicio) & ":AA" & (lnFilaInicio))
            'loFuente = loRango.Font
            'loFuente.Bold = True

            '****************************************************************************************
            ' Ajustes finales de formato (tamaño de celdas, etc...)
            '****************************************************************************************
            loFilas = loCeldas.Rows
            loFilas.AutoFit()

            loColumnas = loCeldas.Columns
            loColumnas.AutoFit()

            loRango = loHoja.Range("A1:A" & lnFilaInicio)
            loRango.ColumnWidth = 2

            loRango = loHoja.Range("B1:B" & lnFilaInicio)
            loRango.ColumnWidth = 6

            loRango = loHoja.Range("C1:C" & lnFilaInicio)
            loRango.ColumnWidth = 11

            loRango = loHoja.Range("D1:D" & lnFilaInicio)
            loRango.ColumnWidth = 11

            loRango = loHoja.Range("E1:E" & lnFilaInicio)
            loRango.ColumnWidth = 14

            loRango = loHoja.Range("F1:F" & lnFilaInicio)
            loRango.ColumnWidth = 35

            loRango = loHoja.Range("G1:G" & lnFilaInicio)
            loRango.ColumnWidth = 18

            loRango = loHoja.Range("H1:H" & lnFilaInicio)
            loRango.ColumnWidth = 13

            loRango = loHoja.Range("I1:I" & lnFilaInicio)
            loRango.ColumnWidth = 13

            loRango = loHoja.Range("J1:J" & lnFilaInicio)
            loRango.ColumnWidth = 16

            loRango = loHoja.Range("K1:K" & lnFilaInicio)
            loRango.ColumnWidth = 13

            loRango = loHoja.Range("L1:L" & lnFilaInicio)
            loRango.ColumnWidth = 13

            loRango = loHoja.Range("M1:M" & lnFilaInicio)
            loRango.ColumnWidth = 10

            loRango = loHoja.Range("N1:N" & lnFilaInicio)
            loRango.ColumnWidth = 13

            loRango = loHoja.Range("O1:T" & lnFilaInicio)
            loRango.ColumnWidth = 14

            loRango = loHoja.Range("U1:U" & lnFilaInicio)
            loRango.ColumnWidth = 11

            loRango = loHoja.Range("V1:X" & lnFilaInicio)
            loRango.ColumnWidth = 14

            loRango = loHoja.Range("Y1:Y" & lnFilaInicio)
            loRango.ColumnWidth = 11

            loRango = loHoja.Range("Z1:AA" & lnFilaInicio)
            loRango.ColumnWidth = 14

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' EAG: 06/10/15: Codigo inicial.					                                        '
'-------------------------------------------------------------------------------------------'
