'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Ventas_SID"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Ventas_SID
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = cusAplicacion.goReportes.paParametrosFinales(5)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpLibroVentas(	Documento	    VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Control		    VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Fec_Ini		    DATETIME,")
            loConsulta.AppendLine("								Nom_Cli		    VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Rif			    VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Cod_Tip		    VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Doc_Ori		    VARCHAR(13) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Mon_Des 	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Rec 	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Por_Des 	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Por_Rec		    DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Otr1	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Otr2	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Otr3	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Net 	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Bru 	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Exe 	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Cod_Imp 	    VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Por_Imp1 	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Imp1 	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Com_Ret 	    VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								Fec_Ret 	    DATETIME,")
            loConsulta.AppendLine("								Mon_Ret 	    DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Net_Anu 	DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Bru_Anu 	DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Exe_Anu 	DECIMAL(28,10),")
            loConsulta.AppendLine("								Por_Imp1_Anu 	DECIMAL(28,10),")
            loConsulta.AppendLine("								Mon_Imp1_Anu 	DECIMAL(28,10))")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpLibroVentas(Documento, Control, Fec_Ini, Nom_Cli, Rif, Cod_Tip, Doc_Ori,")
            loConsulta.AppendLine("							Mon_Des, Mon_Rec, Por_Des, Por_Rec, Mon_Otr1, Mon_Otr2, Mon_Otr3,")
            loConsulta.AppendLine("							Mon_Net, Mon_Bru, Mon_Exe, Cod_Imp, Por_Imp1, Mon_Imp1)")
            loConsulta.AppendLine("SELECT		Cuentas_Cobrar.Documento								AS Documento,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Control									AS Control,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Fec_Ini									AS Fec_Ini, ")
            loConsulta.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Nom_Cli = ''")
            loConsulta.AppendLine("					THEN Clientes.Nom_Cli ")
            loConsulta.AppendLine("					ELSE Cuentas_Cobrar.Nom_Cli")
            loConsulta.AppendLine("			END)													AS Nom_Cli,")
            loConsulta.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Rif = ''")
            loConsulta.AppendLine("					THEN Clientes.Rif ")
            loConsulta.AppendLine("					ELSE Cuentas_Cobrar.Rif")
            loConsulta.AppendLine("			END)													AS Rif,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Cod_Tip									AS Cod_Tip,")
            loConsulta.AppendLine("			(CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loConsulta.AppendLine("					THEN '***ANULADA***' ELSE ''")
            loConsulta.AppendLine("			END)													AS Doc_Ori, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Des  								AS Mon_Des,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Rec  								AS Mon_Rec,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Por_Des  								AS Por_Des,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Por_Rec  								AS Por_Rec,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Otr1 								AS Mon_Otr1,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Otr2 								AS Mon_Otr2,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Otr3 								AS Mon_Otr3,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Net  								AS Mon_Net, ")
            loConsulta.AppendLine("			(Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe)   	AS Mon_Bru, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Exe  								AS Mon_Exe, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Cod_Imp									AS Cod_Imp, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Por_Imp1 								AS Por_Imp1, ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Imp1 								AS Mon_Imp1")
            loConsulta.AppendLine("FROM		Cuentas_Cobrar")
            loConsulta.AppendLine("	JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli   =   Clientes.Cod_Cli")
            loConsulta.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip      =   'FACT' ")
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("			AND " & lcParametro0Hasta)
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("			AND " & lcParametro1Hasta)
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("			AND " & lcParametro2Hasta)
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("			AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
            
        '*****************	Facturas de Venta anuladas (no generan CxC) *************************************  
 
            loConsulta.AppendLine("")
		    loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT		Facturas.Documento								AS Documento,")
            loConsulta.AppendLine("			Facturas.Control								AS Control,")
            loConsulta.AppendLine("			Facturas.Fec_Ini								AS Fec_Ini, ")
            loConsulta.AppendLine("			(CASE	WHEN Facturas.Nom_Cli = ''")
            loConsulta.AppendLine("					THEN Clientes.Nom_Cli ")
            loConsulta.AppendLine("					ELSE Facturas.Nom_Cli")
            loConsulta.AppendLine("			END)											AS Nom_Cli,")
            loConsulta.AppendLine("			(CASE	WHEN Facturas.Rif = ''")
            loConsulta.AppendLine("					THEN Clientes.Rif ")
            loConsulta.AppendLine("					ELSE Facturas.Rif")
            loConsulta.AppendLine("			END)											AS Rif,")
            loConsulta.AppendLine("			'FACT'      									AS Cod_Tip, ")
            loConsulta.AppendLine("			'***ANULADA***'              					AS Doc_Ori, ")
            loConsulta.AppendLine("			Facturas.mon_des1  								AS Mon_Des, ")
            loConsulta.AppendLine("			Facturas.mon_rec1  								AS Mon_Rec, ")
            loConsulta.AppendLine("			Facturas.Por_Des1  								AS Por_Des, ")
            loConsulta.AppendLine("			Facturas.Por_Rec1  								AS Por_Rec, ")
            loConsulta.AppendLine("			Facturas.Mon_Otr1 								AS Mon_Otr1,")
            loConsulta.AppendLine("			Facturas.Mon_Otr2 								AS Mon_Otr2,")
            loConsulta.AppendLine("			Facturas.Mon_Otr3 								AS Mon_Otr3,")
            loConsulta.AppendLine("			Facturas.Mon_Net								AS Mon_Net, ")
            loConsulta.AppendLine("			Facturas.Mon_Bru								AS Mon_Bru, ")
            loConsulta.AppendLine("			Facturas.Mon_Exe								AS Mon_Exe, ")
            loConsulta.AppendLine("			Facturas.cod_imp1								AS Cod_Imp, ")
            loConsulta.AppendLine("			Facturas.Por_Imp1								AS Por_Imp1,")
            loConsulta.AppendLine("			Facturas.Mon_Imp1								AS Mon_Imp1 ")
            loConsulta.AppendLine("FROM		Facturas")
            loConsulta.AppendLine("	JOIN	Clientes ON Facturas.Cod_Cli   =   Clientes.Cod_Cli")
            loConsulta.AppendLine("WHERE		Facturas.Status      =   'Anulado' ")
            loConsulta.AppendLine("			AND Facturas.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("			AND " & lcParametro0Hasta)
            loConsulta.AppendLine("			AND Facturas.Documento    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("			AND " & lcParametro1Hasta)
            loConsulta.AppendLine("			AND Facturas.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("			AND " & lcParametro2Hasta)
            loConsulta.AppendLine("			AND Facturas.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("			AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Facturas.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Facturas.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            
        '*****************	Notas de Débito *************************************  
            loConsulta.AppendLine("")
			loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT		Cuentas_Cobrar.Documento							AS Documento,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Control								AS Control,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Fec_Ini								AS Fec_Ini, ")
            loConsulta.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Nom_Cli = ''")
            loConsulta.AppendLine("					THEN Clientes.Nom_Cli ")
            loConsulta.AppendLine("					ELSE Cuentas_Cobrar.Nom_Cli")
            loConsulta.AppendLine("			END)												AS Nom_Cli,")
            loConsulta.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Rif = ''")
            loConsulta.AppendLine("					THEN Clientes.Rif ")
            loConsulta.AppendLine("					ELSE Cuentas_Cobrar.Rif")
            loConsulta.AppendLine("			END)												AS Rif,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Cod_Tip								AS Cod_Tip,")
            loConsulta.AppendLine("			(CASE WHEN Cuentas_Cobrar.Status = 'Anulado'")
            loConsulta.AppendLine("					THEN '***ANULADA***' ELSE Cuentas_Cobrar.Doc_Ori")
            loConsulta.AppendLine("			END)												AS Doc_Ori,     ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Des  							AS Mon_Des,     ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Rec  							AS Mon_Rec,     ")
            loConsulta.AppendLine("       	Cuentas_Cobrar.Por_Des  							AS Por_Des,     ")
            loConsulta.AppendLine("       	Cuentas_Cobrar.Por_Rec  							AS Por_Rec,     ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Otr1 							AS Mon_Otr1,    ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Otr2 							AS Mon_Otr2,    ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Otr3 							AS Mon_Otr3,    ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Net 							    AS Mon_Net,     ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Bru                              AS Mon_Bru,     ")
            loConsulta.AppendLine("		    Cuentas_Cobrar.Mon_Exe  							AS Mon_Exe,     ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Cod_Imp								AS Cod_Imp,     ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Por_Imp1 							AS Por_Imp1,    ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Imp1 							AS Mon_Imp1")
            loConsulta.AppendLine("FROM		Cuentas_Cobrar")
            loConsulta.AppendLine("	JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli   =   Clientes.Cod_Cli")
            loConsulta.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip      =   'N/DB' ")
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("			AND " & lcParametro0Hasta)
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("			AND " & lcParametro1Hasta)
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("			AND " & lcParametro2Hasta)
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("			AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("")
            
        '*****************	Notas de Crédito *************************************  
            loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("SELECT		Cuentas_Cobrar.Documento							AS Documento,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Control								AS Control,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Fec_Ini								AS Fec_Ini, ")
            loConsulta.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Nom_Cli = ''")
            loConsulta.AppendLine("					THEN Clientes.Nom_Cli ")
            loConsulta.AppendLine("					ELSE Cuentas_Cobrar.Nom_Cli")
            loConsulta.AppendLine("			END)												AS Nom_Cli,")
            loConsulta.AppendLine("			(CASE	WHEN Cuentas_Cobrar.Rif = ''")
            loConsulta.AppendLine("					THEN Clientes.Rif ")
            loConsulta.AppendLine("					ELSE Cuentas_Cobrar.Rif")
            loConsulta.AppendLine("			END)												AS Rif,")
            loConsulta.AppendLine("			Cuentas_Cobrar.Cod_Tip								AS Cod_Tip,")
            loConsulta.AppendLine("			(CASE WHEN Cuentas_Cobrar.Status = 'Anulado'                        ")
            loConsulta.AppendLine("					THEN '***ANULADA***'                                        ")
            loConsulta.AppendLine("					ELSE (CASE WHEN Cuentas_Cobrar.Tip_Ori = 'Devoluciones_Clientes'")
            loConsulta.AppendLine("					    THEN Cuentas_Cobrar.Doc_Ori                             ")
            loConsulta.AppendLine("					    ELSE ''                                                 ")
            loConsulta.AppendLine("					END)                                                        ")
            loConsulta.AppendLine("			END)												AS Doc_Ori,     ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Des*(-1)  						AS Mon_Des,     ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Rec*(-1)  						AS Mon_Rec,     ")
            loConsulta.AppendLine("       	Cuentas_Cobrar.Por_Des  							AS Por_Des,     ")
            loConsulta.AppendLine("       	Cuentas_Cobrar.Por_Rec  							AS Por_Rec,     ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Otr1*(-1) 						AS Mon_Otr1,    ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Otr2*(-1) 						AS Mon_Otr2,    ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Otr3*(-1) 						AS Mon_Otr3,    ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Net*(-1)                         AS Mon_Net,     ")
            loConsulta.AppendLine("			(Cuentas_Cobrar.Mon_Bru - Cuentas_Cobrar.Mon_Exe)*(-1) AS Mon_Bru,  ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Exe*(-1) 						AS Mon_Exe,     ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Cod_Imp								AS Cod_Imp,     ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Por_Imp1 							AS Por_Imp1,    ")
            loConsulta.AppendLine("			Cuentas_Cobrar.Mon_Imp1*(-1)						AS Mon_Imp1     ")
            loConsulta.AppendLine("FROM		Cuentas_Cobrar")
            loConsulta.AppendLine("	JOIN	Clientes ON Cuentas_Cobrar.Cod_Cli   =   Clientes.Cod_Cli")
            loConsulta.AppendLine("WHERE		Cuentas_Cobrar.Cod_Tip      =   'N/CR' ")
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Fec_Ini      BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("			AND " & lcParametro0Hasta)
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Documento    BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("			AND " & lcParametro1Hasta)
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Cod_Cli      BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("			AND " & lcParametro2Hasta)
            loConsulta.AppendLine("			AND Cuentas_Cobrar.Cod_Suc      BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("			AND " & lcParametro3Hasta)
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Cobrar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("")
            
        '*****************	Oculta los montos anulados *************************************  

            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE		#tmpLibroVentas")
            loConsulta.AppendLine("SET         Mon_Net_Anu  = 0,")
            loConsulta.AppendLine("            Mon_Net      = 0,")
            loConsulta.AppendLine("            Mon_Bru_Anu  = 0,")
            loConsulta.AppendLine("            Mon_Bru      = 0,")
            loConsulta.AppendLine("            Mon_Exe_Anu  = 0,")
            loConsulta.AppendLine("            Mon_Exe      = 0,")
            loConsulta.AppendLine("            Por_Imp1_Anu = 0,")
            loConsulta.AppendLine("            Por_Imp1     = 0,")
            loConsulta.AppendLine("            Mon_Imp1_Anu = 0,")
            loConsulta.AppendLine("            Mon_Imp1     = 0,")
            loConsulta.AppendLine("            Nom_Cli      = '***ANULADA***',")
            loConsulta.AppendLine("            Rif          = '',")
            loConsulta.AppendLine("            Doc_Ori      = ''")
            loConsulta.AppendLine("WHERE       Doc_Ori      = '***ANULADA***'")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

        '*****************	Facturas afectadas por las devoluciones *************************************  

            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE      #tmpLibroVentas")
            loConsulta.AppendLine("SET         Doc_Ori = COALESCE((SELECT  TOP 1 Doc_Ori ")
            loConsulta.AppendLine("                                FROM    renglones_dclientes ")
            loConsulta.AppendLine("                                WHERE   documento = #tmpLibroVentas.Doc_Ori ")
            loConsulta.AppendLine("                                    AND tip_ori='Facturas'), '')")
            loConsulta.AppendLine("WHERE       Doc_Ori <> '' AND Cod_Tip = 'N/CR' ")
            loConsulta.AppendLine("")

        '*****************	Retenciones de IVA, si aplican *************************************  
			loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE		#tmpLibroVentas")
            loConsulta.AppendLine("SET			Com_Ret = Retenciones.Com_Ret,")
            loConsulta.AppendLine("			Fec_Ret = Retenciones.Fec_Ret,")
            loConsulta.AppendLine("			Mon_Ret = Retenciones.Mon_Ret")
            loConsulta.AppendLine("FROM	(	SELECT		#tmpLibroVentas.Documento		AS Documento,")
            loConsulta.AppendLine("						#tmpLibroVentas.Cod_Tip			AS Cod_Tip,")
            loConsulta.AppendLine("						retenciones_documentos.num_com	AS Com_Ret, ")
            loConsulta.AppendLine("						cuentas_cobrar.fec_ini			AS Fec_Ret, ")
            loConsulta.AppendLine("						cuentas_cobrar.mon_net			AS Mon_Ret")
            loConsulta.AppendLine("			FROM		retenciones_documentos ")
            loConsulta.AppendLine("				JOIN	#tmpLibroVentas")
            loConsulta.AppendLine("					ON	#tmpLibroVentas.Documento = retenciones_documentos.doc_ori")
            loConsulta.AppendLine("					AND	#tmpLibroVentas.Cod_Tip = retenciones_documentos.cla_ori")
            loConsulta.AppendLine("					AND	retenciones_documentos.tip_ori = 'Cuentas_Cobrar'")
            loConsulta.AppendLine("					AND	retenciones_documentos.Clase = 'IMPUESTO'")
            loConsulta.AppendLine("				JOIN	cuentas_cobrar ")
            loConsulta.AppendLine("					ON	cuentas_cobrar.documento = retenciones_documentos.doc_des")
            loConsulta.AppendLine("					AND	cuentas_cobrar.cod_tip = retenciones_documentos.cla_des")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("		) AS Retenciones")
            loConsulta.AppendLine("WHERE	Retenciones.Documento = #tmpLibroVentas.documento")
            loConsulta.AppendLine("	AND	Retenciones.Cod_Tip = #tmpLibroVentas.Cod_Tip")
            loConsulta.AppendLine("")
            
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Obtiene los impuestos")
            loConsulta.AppendLine("DECLARE @lcImpuesto AS VARCHAR(30);")
            loConsulta.AppendLine("SET @lcImpuesto = '';")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT @lcImpuesto = @lcImpuesto + CONVERT(VARCHAR(20), CAST(Por_Imp1 AS DECIMAL(8,2)), 2) + '%; '")
            loConsulta.AppendLine("FROM #tmpLibroVentas")
            loConsulta.AppendLine("WHERE Por_Imp1>0")
            loConsulta.AppendLine("GROUP BY Por_Imp1")
            loConsulta.AppendLine("ORDER BY Por_Imp1 DESC;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("IF(LEN(@lcImpuesto)>0)")
            loConsulta.AppendLine("    SELECT @lcImpuesto = '(' + SUBSTRING(RTRIM(REPLACE(@lcImpuesto, '.', ',')), 1, LEN(@lcImpuesto)-1) + ')';")
            loConsulta.AppendLine("ELSE")
            loConsulta.AppendLine("    SELECT @lcImpuesto ='';")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- SELECT Final")
            loConsulta.AppendLine("SELECT	Documento, Control, Fec_Ini, Nom_Cli, Rif, Cod_Tip, Doc_Ori,")
            loConsulta.AppendLine("		Mon_Des, Mon_Rec, Por_Des, Por_Rec, Mon_Otr1, Mon_Otr2, Mon_Otr3,")
            loConsulta.AppendLine("		Mon_Net, Mon_Bru, Mon_Exe, Cod_Imp, Por_Imp1, Mon_Imp1,")
            loConsulta.AppendLine("		Com_Ret, Fec_Ret, Mon_Ret,")
            loConsulta.AppendLine("        Mon_Net_Anu, Mon_Bru_Anu, Mon_Exe_Anu, ")
            loConsulta.AppendLine("        Por_Imp1_Anu, Mon_Imp1_Anu, @lcImpuesto AS Impuestos")
            loConsulta.AppendLine("FROM	#tmpLibroVentas")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Ventas_SID", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLibro_Ventas_SID.ReportSource = loObjetoReporte

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
' RJG: 03/09/13: Codigo inicial, a partir de rLibro_Ventas_2013.aspx.						'
'                Se ocultaron los datos de las facturas anuladas (cliente, montos...). Se   '
'                ajustó el título. '
'-------------------------------------------------------------------------------------------'
' RJG: 03/10/13: Se movió el texto "***ANULADA***" a la columna del cliente, se agregaron   '
'                los totales, se corrigió el número de factura afectada, se ajusto el orden.'
'-------------------------------------------------------------------------------------------'
