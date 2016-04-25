'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rLibro_Ventas"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rLibro_Ventas
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpLibroVentas(	Documento	    VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Control		    VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Fec_Ini		    DATETIME,")
            loComandoSeleccionar.AppendLine("								Nom_Cli		    VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Rif			    VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Cod_Tip		    VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Doc_Ori		    VARCHAR(13) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Mon_Des 	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Rec 	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Por_Des 	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Por_Rec		    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Otr1	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Otr2	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Otr3	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Net 	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Bru 	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Exe 	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Cod_Imp 	    VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Por_Imp1 	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Imp1 	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Com_Ret 	    VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Fec_Ret 	    DATETIME,")
            loComandoSeleccionar.AppendLine("								Mon_Ret 	    DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Net_Anu 	DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Bru_Anu 	DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Exe_Anu 	DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Por_Imp1_Anu 	DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Imp1_Anu 	DECIMAL(28,10))")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpLibroVentas(Documento, Control, Fec_Ini, Nom_Cli, Rif, Cod_Tip, Doc_Ori,")
            loComandoSeleccionar.AppendLine("							Mon_Des, Mon_Rec, Por_Des, Por_Rec, Mon_Otr1, Mon_Otr2, Mon_Otr3,")
            loComandoSeleccionar.AppendLine("							Mon_Net, Mon_Bru, Mon_Exe, Cod_Imp, Por_Imp1, Mon_Imp1)")
            loComandoSeleccionar.AppendLine("SELECT		Cuentas_Cobrar.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Control									AS Control,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini									AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("					THEN 'ANULADO' ")
            loComandoSeleccionar.AppendLine("					ELSE (CASE	WHEN Cuentas_Cobrar.Nom_Cli = ''")
            loComandoSeleccionar.AppendLine("                               THEN  Clientes.Nom_Cli")
            loComandoSeleccionar.AppendLine("                               ELSE Cuentas_Cobrar.Nom_Cli END)")
            loComandoSeleccionar.AppendLine("			END)													AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Rif = ''")
            loComandoSeleccionar.AppendLine("					THEN Clientes.Rif ")
            loComandoSeleccionar.AppendLine("					ELSE Cuentas_Cobrar.Rif")
            loComandoSeleccionar.AppendLine("			END)													AS Rif,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Tip									AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("				  THEN '' ELSE Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("			END)													AS Doc_Ori, ")
            loComandoSeleccionar.AppendLine("			CASE	WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("			        THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("			        ELSE Cuentas_Cobrar.Mon_Des  END				AS Mon_Des,")
            loComandoSeleccionar.AppendLine("			CASE	WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("			        THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("			        ELSE Cuentas_Cobrar.Mon_Rec  END				AS Mon_Rec,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Des  								AS Por_Des,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Rec  								AS Por_Rec,")
            loComandoSeleccionar.AppendLine("			CASE	WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("			        THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("			        ELSE Cuentas_Cobrar.Mon_Otr1 END				AS Mon_Otr1,")
            loComandoSeleccionar.AppendLine("			CASE	WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("			        THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("			        ELSE Cuentas_Cobrar.Mon_Otr2 END				AS Mon_Otr2,")
            loComandoSeleccionar.AppendLine("			CASE	WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("			        THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("			        ELSE Cuentas_Cobrar.Mon_Otr3 END				AS Mon_Otr3,")
            loComandoSeleccionar.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("			        THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("			        ELSE Cuentas_Cobrar.Mon_Net END) 				AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			CASE	WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("			        THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("			        ELSE (Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe) END	AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			CASE	WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("			        THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("			        ELSE Cuentas_Cobrar.Mon_Exe END					AS Mon_Exe, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Imp									AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Imp1 								AS Por_Imp1, ")
            loComandoSeleccionar.AppendLine("			CASE	WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("			        THEN CAST('0.00' AS DECIMAL(28,2))")
            loComandoSeleccionar.AppendLine("			        ELSE Cuentas_Cobrar.Mon_Imp1 END				AS Mon_Imp1")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli   =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip      =   'FACT' ")
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            

            '*****************	Facturas de Venta anuladas (no generan CxC) *************************************  

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Facturas.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("			Facturas.Control								AS Control,")
            loComandoSeleccionar.AppendLine("			Facturas.Fec_Ini								AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			'ANULADO'	           							AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("			(CASE	WHEN Facturas.Rif = ''")
            loComandoSeleccionar.AppendLine("					THEN Clientes.Rif ")
            loComandoSeleccionar.AppendLine("					ELSE Facturas.Rif")
            loComandoSeleccionar.AppendLine("			END)											AS Rif,")
            loComandoSeleccionar.AppendLine("			'FACT'      									AS Cod_Tip, ")
            loComandoSeleccionar.AppendLine("			''                          					AS Doc_Ori, ")
            loComandoSeleccionar.AppendLine("			Facturas.mon_des1  								AS Mon_Des, ")
            loComandoSeleccionar.AppendLine("			Facturas.mon_rec1  								AS Mon_Rec, ")
            loComandoSeleccionar.AppendLine("			Facturas.Por_Des1  								AS Por_Des, ")
            loComandoSeleccionar.AppendLine("			Facturas.Por_Rec1  								AS Por_Rec, ")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Otr1 								AS Mon_Otr1,")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Otr2 								AS Mon_Otr2,")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Otr3 								AS Mon_Otr3,")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Net								AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Bru								AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Exe								AS Mon_Exe, ")
            loComandoSeleccionar.AppendLine("			Facturas.cod_imp1								AS Cod_Imp, ")
            loComandoSeleccionar.AppendLine("			Facturas.Por_Imp1								AS Por_Imp1,")
            loComandoSeleccionar.AppendLine("			Facturas.Mon_Imp1								AS Mon_Imp1 ")
            loComandoSeleccionar.AppendLine("FROM		Facturas")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON Facturas.Cod_Cli   =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE		Facturas.Status      =   'Anulado' ")
            loComandoSeleccionar.AppendLine("			AND Facturas.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            '*****************	Notas de Débito *************************************  
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Cuentas_Cobrar.Documento							AS Documento,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Control								AS Control,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini								AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Nom_Cli = ''")
            loComandoSeleccionar.AppendLine("					THEN Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine("					ELSE Cuentas_Cobrar.Nom_Cli")
            loComandoSeleccionar.AppendLine("			END)												AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Rif = ''")
            loComandoSeleccionar.AppendLine("					THEN Clientes.Rif ")
            loComandoSeleccionar.AppendLine("					ELSE Cuentas_Cobrar.Rif")
            loComandoSeleccionar.AppendLine("			END)												AS Rif,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Tip								AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("					THEN '***ANULADA***' ELSE Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("			END)												AS Doc_Ori,     ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Des  							AS Mon_Des,     ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Rec  							AS Mon_Rec,     ")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Por_Des  							AS Por_Des,     ")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Por_Rec  							AS Por_Rec,     ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr1 							AS Mon_Otr1,    ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr2 							AS Mon_Otr2,    ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr3 							AS Mon_Otr3,    ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Net 							    AS Mon_Net,     ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Bru                              AS Mon_Bru,     ")
            loComandoSeleccionar.AppendLine("		    Cuentas_Cobrar.Mon_Exe  							AS Mon_Exe,     ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Imp								AS Cod_Imp,     ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Imp1 							AS Por_Imp1,    ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Imp1 							AS Mon_Imp1")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli   =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip      =   'N/DB' ")
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            
            loComandoSeleccionar.AppendLine("")

            '*****************	Notas de Crédito *************************************  
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("SELECT		Cuentas_Cobrar.Documento							AS Documento,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Control								AS Control,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Fec_Ini								AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Nom_Cli = ''")
            loComandoSeleccionar.AppendLine("					THEN Clientes.Nom_Cli ")
            loComandoSeleccionar.AppendLine("					ELSE Cuentas_Cobrar.Nom_Cli")
            loComandoSeleccionar.AppendLine("			END)												AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Rif = ''")
            loComandoSeleccionar.AppendLine("					THEN Clientes.Rif ")
            loComandoSeleccionar.AppendLine("					ELSE Cuentas_Cobrar.Rif")
            loComandoSeleccionar.AppendLine("			END)												AS Rif,")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Tip								AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loComandoSeleccionar.AppendLine("					THEN '***ANULADA***' ELSE Cuentas_Cobrar.Referencia")
            loComandoSeleccionar.AppendLine("			END)												AS Doc_Ori,     ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Des*(-1)  						AS Mon_Des,     ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Rec*(-1)  						AS Mon_Rec,     ")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Por_Des  							AS Por_Des,     ")
            loComandoSeleccionar.AppendLine("       	Cuentas_Cobrar.Por_Rec  							AS Por_Rec,     ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr1*(-1) 						AS Mon_Otr1,    ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr2*(-1) 						AS Mon_Otr2,    ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Otr3*(-1) 						AS Mon_Otr3,    ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Net*(-1)                         AS Mon_Net,     ")
            loComandoSeleccionar.AppendLine("			(Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe)*(-1) AS Mon_Bru,  ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Exe*(-1) 						AS Mon_Exe,     ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Imp								AS Cod_Imp,     ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Por_Imp1 							AS Por_Imp1,    ")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Imp1*(-1)						AS Mon_Imp1     ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli   =   Clientes.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip      =   'N/CR' ")
            loComandoSeleccionar.AppendLine("			AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
           
            loComandoSeleccionar.AppendLine("")

            '*****************	Oculta los montos anulados *************************************  

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE		#tmpLibroVentas")
            loComandoSeleccionar.AppendLine("SET         Mon_Net_Anu = Mon_Net,")
            loComandoSeleccionar.AppendLine("            Mon_Net = NULL,")
            loComandoSeleccionar.AppendLine("            Mon_Bru_Anu = Mon_Bru,")
            loComandoSeleccionar.AppendLine("            Mon_Bru = NULL,")
            loComandoSeleccionar.AppendLine("            Mon_Exe_Anu = Mon_Exe,")
            loComandoSeleccionar.AppendLine("            Mon_Exe = NULL,")
            loComandoSeleccionar.AppendLine("            Por_Imp1_Anu = Por_Imp1,")
            loComandoSeleccionar.AppendLine("            Por_Imp1 = NULL,")
            loComandoSeleccionar.AppendLine("            Mon_Imp1_Anu = Mon_Imp1,")
            loComandoSeleccionar.AppendLine("            Mon_Imp1 = NULL")
            loComandoSeleccionar.AppendLine("WHERE       Doc_Ori = '***ANULADA***'")
            loComandoSeleccionar.AppendLine("")

            '*****************	Retenciones de IVA, si aplican *************************************  
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UPDATE		#tmpLibroVentas")
            loComandoSeleccionar.AppendLine("SET			Com_Ret = Retenciones.Com_Ret,")
            loComandoSeleccionar.AppendLine("			Fec_Ret = Retenciones.Fec_Ret,")
            loComandoSeleccionar.AppendLine("			Mon_Ret = Retenciones.Mon_Ret")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT		#tmpLibroVentas.Documento		AS Documento,")
            loComandoSeleccionar.AppendLine("						#tmpLibroVentas.Cod_Tip			AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("						retenciones_documentos.num_com	AS Com_Ret, ")
            loComandoSeleccionar.AppendLine("						cuentas_cobrar.fec_ini			AS Fec_Ret, ")
            loComandoSeleccionar.AppendLine("						cuentas_cobrar.mon_net			AS Mon_Ret")
            loComandoSeleccionar.AppendLine("			FROM		retenciones_documentos ")
            loComandoSeleccionar.AppendLine("				JOIN	#tmpLibroVentas")
            loComandoSeleccionar.AppendLine("					ON	#tmpLibroVentas.Documento = retenciones_documentos.doc_ori")
            loComandoSeleccionar.AppendLine("					AND	#tmpLibroVentas.Cod_Tip = retenciones_documentos.cla_ori")
            loComandoSeleccionar.AppendLine("					AND	retenciones_documentos.tip_ori = 'Cuentas_Cobrar'")
            loComandoSeleccionar.AppendLine("					AND	retenciones_documentos.Clase = 'IMPUESTO'")
            loComandoSeleccionar.AppendLine("				JOIN	cuentas_cobrar ")
            loComandoSeleccionar.AppendLine("					ON	cuentas_cobrar.documento = retenciones_documentos.doc_des")
            loComandoSeleccionar.AppendLine("					AND	cuentas_cobrar.cod_tip = retenciones_documentos.cla_des")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("		) AS Retenciones")
            loComandoSeleccionar.AppendLine("WHERE	Retenciones.Documento = #tmpLibroVentas.documento")
            loComandoSeleccionar.AppendLine("	AND	Retenciones.Cod_Tip = #tmpLibroVentas.Cod_Tip")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- Obtiene los impuestos")
            loComandoSeleccionar.AppendLine("DECLARE @lcImpuesto AS VARCHAR(30);")
            loComandoSeleccionar.AppendLine("SET @lcImpuesto = '';")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT @lcImpuesto = @lcImpuesto + CONVERT(VARCHAR(20), CAST(Por_Imp1 AS DECIMAL(8,2)), 2) + '%; '")
            loComandoSeleccionar.AppendLine("FROM #tmpLibroVentas")
            loComandoSeleccionar.AppendLine("WHERE Por_Imp1>0")
            loComandoSeleccionar.AppendLine("GROUP BY Por_Imp1")
            loComandoSeleccionar.AppendLine("ORDER BY Por_Imp1 DESC;")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("IF(LEN(@lcImpuesto)>0)")
            loComandoSeleccionar.AppendLine("    SELECT @lcImpuesto = '(' + SUBSTRING(RTRIM(REPLACE(@lcImpuesto, '.', ',')), 1, LEN(@lcImpuesto)-1) + ')';")
            loComandoSeleccionar.AppendLine("ELSE")
            loComandoSeleccionar.AppendLine("    SELECT @lcImpuesto ='';")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- SELECT Final")
            loComandoSeleccionar.AppendLine("SELECT	ROW_NUMBER() OVER (ORDER BY Registros.Fec_Ini) AS Num,*")
            loComandoSeleccionar.AppendLine("FROM(SELECT	Documento, Control, Fec_Ini, Nom_Cli, Rif, Cod_Tip, Doc_Ori,")
            loComandoSeleccionar.AppendLine("		Mon_Des, Mon_Rec, Por_Des, Por_Rec, Mon_Otr1, Mon_Otr2, Mon_Otr3,")
            loComandoSeleccionar.AppendLine("		Mon_Net, Mon_Bru, Mon_Exe, Cod_Imp, Por_Imp1, Mon_Imp1,")
            loComandoSeleccionar.AppendLine("		Com_Ret, Fec_Ret, Mon_Ret,")
            loComandoSeleccionar.AppendLine("        Mon_Net_Anu, Mon_Bru_Anu, Mon_Exe_Anu, ")
            loComandoSeleccionar.AppendLine("        Por_Imp1_Anu, Mon_Imp1_Anu, @lcImpuesto AS Impuestos,")
            loComandoSeleccionar.AppendLine("       MONTH(" & lcParametro0Desde & " )				AS Mes,")
            loComandoSeleccionar.AppendLine("       YEAR(" & lcParametro0Hasta & " )				AS Anio")
            loComandoSeleccionar.AppendLine("FROM	#tmpLibroVentas) Registros")
            'loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpLibroVentas")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")



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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rLibro_Ventas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rLibro_Ventas.ReportSource = loObjetoReporte

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
' RJG: 28/02/13: Codigo inicial, a partir de rLibro_Ventas.aspx.							'
'-------------------------------------------------------------------------------------------'
' RJG: 16/04/13: Agregado filtro para incluir solo retenciones de IVA (no ISLR ni Patente). '
'-------------------------------------------------------------------------------------------'
' RJG: 29/07/13: Se agregaron las Facturas de Venta Anuladas. También se mostrarán los      '
'                montos de los documetnos anulados, pero sin contarlos para los totales. Se '
'                ajustaron los porcentajes de impuesto en el total para que muestre todos.  '
'-------------------------------------------------------------------------------------------'
