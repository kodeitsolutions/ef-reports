'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Compras_SID"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Compras_SID
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
            Dim lcParametro6Desde As String = cusAplicacion.goReportes.paParametrosFinales(6)
            Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosFinales(7)
            Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE @lnCero AS DECIMAL(28, 10);")
            loConsulta.AppendLine("SET @lnCero = CAST(0 AS DECIMAL(28, 10));")
            loConsulta.AppendLine("DECLARE @lcVacio AS NVARCHAR(30);")
            loConsulta.AppendLine("SET @lcVacio = N'';")

            loConsulta.AppendLine("CREATE TABLE #tmpLibroCompras(	Tabla	                VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Cod_Tip 	            VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Codigo_Tipo	            VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Documento	            VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Control	                VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Referencia	            VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                 Planilla_Importacion    VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Factura 	            VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Status		            VARCHAR(15) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Doc_Ori		            VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Documento_Afectado	    VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Cod_Pro		            VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Fec_Ini		            DATETIME,")
            loConsulta.AppendLine("								    Tip_Doc		            VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Mon_Bru 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    Mon_Net 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    Dis_Imp 	            VARCHAR(MAX),")
            loConsulta.AppendLine("								    Mon_Des 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    Mon_Rec 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    Por_Des 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    Por_Rec 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    Mon_Otr1 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    Mon_Otr2 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    Mon_Otr3 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    Cod_Imp1 	            VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Cod_Imp2 	            VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Cod_Imp3 	            VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Com_Ret		    	    VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Fec_Ret		    	    DATETIME,")
            loConsulta.AppendLine("								    mon_ret 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    mon_imp1 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    mon_bas1 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    por_imp1 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    mon_exe1 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    mon_imp2 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    mon_bas2 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    por_imp2 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    mon_exe2 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    mon_imp3 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    mon_bas3 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    por_imp3 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    mon_exe3 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    subt_exe 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    subt_bas 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    subt_imp 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    Exonerado 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    No_Sujeto 	            DECIMAL(28,10),")
            loConsulta.AppendLine("								    Sin_Derecho_CF          DECIMAL(28,10),")
            loConsulta.AppendLine("								    Nom_Pro 	            VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Rif 	                VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Tip_Pro 	            VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Transaccion		        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("								    Es_Importacion 	        BIT DEFAULT 0,")
            loConsulta.AppendLine("								    Fecha_Inicial		    DATETIME DEFAULT " & lcParametro0Desde & ",")
            loConsulta.AppendLine("								    Fecha_Final             DATETIME DEFAULT " & lcParametro0Hasta & ",")
            loConsulta.AppendLine("								    Impuesto1_EsReducido    BIT DEFAULT 0,")
            loConsulta.AppendLine("								    Impuesto2_EsReducido    BIT DEFAULT 0,")
            loConsulta.AppendLine("								    Impuesto3_EsReducido    BIT DEFAULT 0")
            loConsulta.AppendLine("								)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpLibroCompras(	Tabla, Cod_Tip, Codigo_Tipo, Documento, Control, Referencia, Planilla_Importacion, Factura, Status,")
            loConsulta.AppendLine("								    Doc_Ori, Documento_Afectado, Cod_Pro, Fec_Ini, Tip_Doc, Mon_Bru, Mon_Net,")
            loConsulta.AppendLine("								    Dis_Imp, Mon_Des, Mon_Rec, Por_Des, Por_Rec, Mon_Otr1, Mon_Otr2, Mon_Otr3,")
            loConsulta.AppendLine("								    Cod_Imp1, Cod_Imp2, Cod_Imp3, mon_ret, mon_imp1, mon_bas1, por_imp1,")
            loConsulta.AppendLine("								    mon_exe1, mon_imp2, mon_bas2, por_imp2, mon_exe2, mon_imp3, mon_bas3,")
            loConsulta.AppendLine("								    por_imp3, mon_exe3, subt_exe, subt_bas, subt_imp, Exonerado, No_Sujeto,")
            loConsulta.AppendLine("								    Sin_Derecho_CF, Nom_Pro, Rif, Tip_Pro, Transaccion, Es_Importacion)")
            loConsulta.AppendLine("SELECT		")
            loConsulta.AppendLine("             'Compras'															AS Tabla, 		")
            loConsulta.AppendLine("			    CASE Cuentas_Pagar.Cod_Tip															")
            loConsulta.AppendLine("			     	WHEN 'FACT' 	THEN 'Factura'													")
            loConsulta.AppendLine("			     	WHEN 'N/CR' 	THEN 'Nota de Credito'											")
            loConsulta.AppendLine("			     	WHEN 'N/DB' 	THEN 'Nota de Debito'											")
            loConsulta.AppendLine("			    END																	AS Cod_Tip,		")
            loConsulta.AppendLine("			    Cuentas_Pagar.Cod_Tip												AS Codigo_Tipo,	")
            loConsulta.AppendLine("			    CAST(Cuentas_Pagar.Documento AS CHAR(30))							AS Documento, 	")
            loConsulta.AppendLine("			    Cuentas_Pagar.Control												AS Control, 	")
            loConsulta.AppendLine("			    Cuentas_Pagar.Referencia											AS Referencia,	")
            loConsulta.AppendLine("			    @lcVacio                											AS Planilla_Importacion,	")
            loConsulta.AppendLine("			    Cuentas_Pagar.Factura												AS Factura, 	")
            loConsulta.AppendLine("			    Cuentas_Pagar.Status												AS Status, 		")
            loConsulta.AppendLine("			    Cuentas_Pagar.Doc_Ori												AS Doc_Ori, 	")
            loConsulta.AppendLine("			    @lcVacio															AS Documento_Afectado, 	")
            loConsulta.AppendLine("			    Cuentas_Pagar.Cod_Pro												AS Cod_Pro, 	")
            loConsulta.AppendLine("			    Cuentas_Pagar.Fec_Ini												AS Fec_Ini, 	")
            loConsulta.AppendLine("			    Cuentas_Pagar.Tip_Doc												AS Tip_Doc, 	")
            loConsulta.AppendLine("			    CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("			    	THEN Cuentas_Pagar.Mon_Bru * -1 												")
            loConsulta.AppendLine("			    	ELSE Cuentas_Pagar.Mon_Bru 														")
            loConsulta.AppendLine("			    END																	AS Mon_Bru,  	")
            loConsulta.AppendLine("			    CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("			    	THEN Cuentas_Pagar.Mon_Net * -1 												")
            loConsulta.AppendLine("			    	ELSE Cuentas_Pagar.Mon_Net 														")
            loConsulta.AppendLine("			    END																	AS Mon_Net,  	")
            loConsulta.AppendLine("			    Cuentas_Pagar.Dis_Imp												AS Dis_Imp, 	")
            loConsulta.AppendLine("			    CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("			    	THEN Cuentas_Pagar.Mon_Des * -1 												")
            loConsulta.AppendLine("			    	ELSE Cuentas_Pagar.Mon_Des 														")
            loConsulta.AppendLine("			    END																	AS Mon_Des,  	")
            loConsulta.AppendLine("			    Cuentas_Pagar.Mon_Rec												AS Mon_Rec, 	")
            loConsulta.AppendLine("			    Cuentas_Pagar.Por_Des												AS Por_Des, 	")
            loConsulta.AppendLine("			    Cuentas_Pagar.Por_Rec												AS Por_Rec, 	")
            loConsulta.AppendLine("			    CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("			    	THEN Cuentas_Pagar.Mon_Otr1 * -1 												")
            loConsulta.AppendLine("			    	ELSE Cuentas_Pagar.Mon_Otr1 													")
            loConsulta.AppendLine("			    END																	AS Mon_Otr1,  	")
            loConsulta.AppendLine("			    CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("			    	THEN Cuentas_Pagar.Mon_Otr2 * -1 												")
            loConsulta.AppendLine("			    	ELSE Cuentas_Pagar.Mon_Otr2 													")
            loConsulta.AppendLine("			    END																	AS Mon_Otr2,  	")
            loConsulta.AppendLine("			    CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loConsulta.AppendLine("			    	THEN Cuentas_Pagar.Mon_Otr3 * -1 												")
            loConsulta.AppendLine("			    	ELSE Cuentas_Pagar.Mon_Otr3 													")
            loConsulta.AppendLine("			    END																	AS Mon_Otr3,  	")
            loConsulta.AppendLine("			    @lcVacio															AS Cod_Imp1, 	")
            loConsulta.AppendLine("			    @lcVacio															AS Cod_Imp2, 	")
            loConsulta.AppendLine("			    @lcVacio															AS Cod_Imp3, 	")
            loConsulta.AppendLine("			    @lnCero 															AS mon_ret, 	")
            loConsulta.AppendLine("			    @lnCero 															AS mon_imp1, 	")
            loConsulta.AppendLine("			    @lnCero 															AS mon_bas1, 	")
            loConsulta.AppendLine("			    @lnCero 															AS por_imp1, 	")
            loConsulta.AppendLine("			    @lnCero 															AS mon_exe1, 	")
            loConsulta.AppendLine("			    @lnCero 															AS mon_imp2, 	")
            loConsulta.AppendLine("			    @lnCero 															AS mon_bas2, 	")
            loConsulta.AppendLine("			    @lnCero 															AS por_imp2, 	")
            loConsulta.AppendLine("             @lnCero 															AS mon_exe2, 	")
            loConsulta.AppendLine("             @lnCero 															AS mon_imp3,	")
            loConsulta.AppendLine("             @lnCero 															AS mon_bas3,	")
            loConsulta.AppendLine("             @lnCero 															AS por_imp3,	")
            loConsulta.AppendLine("             @lnCero 															AS mon_exe3,	")
            loConsulta.AppendLine("			    @lnCero 															AS subt_exe,	")
            loConsulta.AppendLine("			    @lnCero 															AS subt_bas,	")
            loConsulta.AppendLine("			    @lnCero 															AS subt_imp,	")
            loConsulta.AppendLine("			    @lnCero 															AS Exonerado, 	")
            loConsulta.AppendLine("			    @lnCero 															AS No_Sujeto, 	")
            loConsulta.AppendLine("			    @lnCero 															AS Sin_Derecho_CF,	")
            loConsulta.AppendLine("			    (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') 				")
            loConsulta.AppendLine("			    	THEN Proveedores.Nom_Pro 														")
            loConsulta.AppendLine("			    	ELSE (CASE WHEN (Cuentas_Pagar.Nom_Pro = '') 									")
            loConsulta.AppendLine("			    		THEN Proveedores.Nom_Pro 													")
            loConsulta.AppendLine("			    		ELSE Cuentas_Pagar.Nom_Pro 													")
            loConsulta.AppendLine("			    	END) 																			")
            loConsulta.AppendLine("			    END)																AS Nom_Pro, 	")
            loConsulta.AppendLine("			    (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') 				")
            loConsulta.AppendLine("			    	THEN Proveedores.Rif ELSE 														")
            loConsulta.AppendLine("			        (CASE WHEN (Cuentas_Pagar.Rif = '')												")
            loConsulta.AppendLine("			    		THEN Proveedores.Rif 														")
            loConsulta.AppendLine("			    		ELSE Cuentas_Pagar.Rif 														")
            loConsulta.AppendLine("			        END) 																			")
            loConsulta.AppendLine("			     END)																AS Rif,			")
            loConsulta.AppendLine("			    Proveedores.Tip_Pro 												AS Tip_Pro,		")
            loConsulta.AppendLine("			    (CASE WHEN Cuentas_Pagar.Status = 'Anulado' 										")
            loConsulta.AppendLine("			    	THEN '03-ANU' 																	")
            loConsulta.AppendLine("			    	ELSE '01-REG' 																	")
            loConsulta.AppendLine("			    END)																AS Transaccion,	")
            loConsulta.AppendLine("			    (CASE WHEN Proveedores.Nacional=1 THEN 0 ELSE 1 END)   				AS Es_Importacion	")
            loConsulta.AppendLine("FROM		    Cuentas_Pagar ")
            loConsulta.AppendLine("	    JOIN	Proveedores ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            loConsulta.AppendLine("WHERE		Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 			AND Cuentas_Pagar.Documento BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro1Hasta)
            loConsulta.AppendLine(" 			AND Cuentas_Pagar.cod_pro BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro2Hasta)
            loConsulta.AppendLine(" 			AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine(" 			AND " & lcParametro3Hasta)
            loConsulta.AppendLine("			    AND Cuentas_Pagar.Status        IN ( " & lcParametro8Desde & " ) ")
            loConsulta.AppendLine("			    AND Cuentas_Pagar.Status <> 'Anulado' ")
            If lcParametro5Desde = "Igual" Then
                loConsulta.AppendLine(" 		AND Cuentas_Pagar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loConsulta.AppendLine(" 		AND Cuentas_Pagar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
            'loConsulta.AppendLine("			AND (										")
            'loConsulta.AppendLine("						Cuentas_Pagar.Cod_Tip IN ('FACT', 'N/DB')")
            'loConsulta.AppendLine("					OR	(Cuentas_Pagar.Cod_Tip = 'N/CR' AND Cuentas_Pagar.Automatico = 1 AND Cuentas_Pagar.Tip_Ori = 'Devoluciones_Proveedores')")
            'loConsulta.AppendLine("					OR	(Cuentas_Pagar.Cod_Tip = 'N/CR' AND Cuentas_Pagar.Automatico = 0 AND Cuentas_Pagar.Cod_Rev IN ('DEVCOM', 'REBAJA') )")
            'loConsulta.AppendLine("				)										")
            loConsulta.AppendLine(" 			AND Cuentas_Pagar.cod_tip IN ('FACT', 'N/CR', 'N/DB')")
            loConsulta.AppendLine("           AND Cuentas_Pagar.Cod_Tip      BETWEEN " & lcParametro9Desde)
            loConsulta.AppendLine("           AND " & lcParametro9Hasta)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim lcStatus As String = "Pendiente,Confirmado,Procesado,Pagado,Cerrado,Afectado,Serializado,Contabilizado,Iniciado,Conciliado,Otro,Anulado"


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''Obtencion de las Ordenes de pago '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If lcParametro6Desde.ToUpper() = "SI" Then


                loConsulta.AppendLine(" UNION ALL")
                loConsulta.AppendLine("")
                loConsulta.AppendLine(" SELECT	")
                loConsulta.AppendLine("			'Orden_Pago'													AS Tabla,		")
                loConsulta.AppendLine("			'Orden de Pago' 												AS cod_tip,		")
                loConsulta.AppendLine("			@lcVacio														AS Codigo_Tipo,	")
                loConsulta.AppendLine("			CAST(Ordenes_Pagos.Documento AS CHAR(30))	 					AS Documento,	")
                loConsulta.AppendLine("			Ordenes_Pagos.Control	 										AS Control,		")
                loConsulta.AppendLine("			@lcVacio														AS Referencia,	")
                loConsulta.AppendLine("			@lcVacio                										AS Planilla_Importacion,		")
                loConsulta.AppendLine("			Ordenes_Pagos.Factura	 										AS Factura,		")
                loConsulta.AppendLine("			Ordenes_Pagos.Status	 										AS Status,		")
                loConsulta.AppendLine("			@lcVacio														AS Doc_Ori, 	")
                loConsulta.AppendLine("			@lcVacio														AS Documento_Afectado, 	")
                loConsulta.AppendLine("			Ordenes_Pagos.Cod_Pro	 										AS Cod_Pro, 	")
                loConsulta.AppendLine("			Ordenes_Pagos.Fec_Ini	 										AS Fec_Ini, 	")
                loConsulta.AppendLine("			'debito'														As Tip_Doc, 	")
                loConsulta.AppendLine("			Ordenes_Pagos.Mon_Bru	 										AS Mon_Bru, 	")
                loConsulta.AppendLine("			Ordenes_Pagos.Mon_Net	 										AS Mon_Net, 	")
                loConsulta.AppendLine("			Ordenes_Pagos.Dis_Imp	 										AS Dis_Imp, 	")
                loConsulta.AppendLine("			@lnCero    				                                        AS Mon_Des, 	")
                loConsulta.AppendLine("			@lnCero    				                                        AS Mon_Rec, 	")
                loConsulta.AppendLine("			@lnCero															AS Por_Des, 	")
                loConsulta.AppendLine("			@lnCero															AS Por_Rec, 	")
                loConsulta.AppendLine("			@lnCero    				                                        AS Mon_Otr1, 	")
                loConsulta.AppendLine("			@lnCero    				                                        AS Mon_Otr2, 	")
                loConsulta.AppendLine("			@lnCero    														AS Mon_Otr3, 	")
                loConsulta.AppendLine("			@lcVacio														AS Cod_Imp1, 	")
                loConsulta.AppendLine("			@lcVacio														AS Cod_Imp2, 	")
                loConsulta.AppendLine("			@lcVacio														AS Cod_Imp3, 	")
                loConsulta.AppendLine("			Ordenes_Pagos.Mon_Ret											AS Mon_Ret,		")
                loConsulta.AppendLine("			@lnCero															AS mon_imp1,	")
                loConsulta.AppendLine("			@lnCero															AS mon_bas1,	")
                loConsulta.AppendLine("			@lnCero															AS por_imp1,	")
                loConsulta.AppendLine("			@lnCero															AS mon_exe1,	")
                loConsulta.AppendLine("			@lnCero															AS mon_imp2,	")
                loConsulta.AppendLine("			@lnCero															AS mon_bas2,	")
                loConsulta.AppendLine("			@lnCero															AS por_imp2,	")
                loConsulta.AppendLine("			@lnCero															AS mon_exe2,	")
                loConsulta.AppendLine("			@lnCero															AS mon_imp3,	")
                loConsulta.AppendLine("			@lnCero															AS mon_bas3,	")
                loConsulta.AppendLine("			@lnCero															AS por_imp3,	")
                loConsulta.AppendLine("			@lnCero															AS mon_exe3,	")
                loConsulta.AppendLine("			@lnCero															AS subt_exe,	")
                loConsulta.AppendLine("			@lnCero															AS subt_bas,	")
                loConsulta.AppendLine("			@lnCero															AS subt_imp,	")
                loConsulta.AppendLine("			@lnCero 														AS Exonerado,	")
                loConsulta.AppendLine("			@lnCero 														AS No_Sujeto,	")
                loConsulta.AppendLine("			@lnCero 														AS Sin_Derecho_CF,	")
                loConsulta.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '')			")
                loConsulta.AppendLine("				THEN Proveedores.Nom_Pro													")
                loConsulta.AppendLine("				ELSE (CASE WHEN (Ordenes_Pagos.Nom_Pro = '')								")
                loConsulta.AppendLine("					THEN Proveedores.Nom_Pro												")
                loConsulta.AppendLine("					ELSE Ordenes_Pagos.Nom_Pro												")
                loConsulta.AppendLine("				END)																		")
                loConsulta.AppendLine("			END)															AS Nom_Pro, 	")
                loConsulta.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '')			")
                loConsulta.AppendLine("				THEN Proveedores.Rif														")
                loConsulta.AppendLine("				ELSE (CASE WHEN (Ordenes_Pagos.Rif = '')									")
                loConsulta.AppendLine("					THEN Proveedores.Rif													")
                loConsulta.AppendLine("					ELSE Ordenes_Pagos.Rif													")
                loConsulta.AppendLine("				END)																		")
                loConsulta.AppendLine("			END)															AS Rif,			")
                loConsulta.AppendLine("			Proveedores.Tip_Pro												AS Tip_Pro,		")
                loConsulta.AppendLine("			(CASE WHEN Ordenes_Pagos.Status = 'Anulado'										")
                loConsulta.AppendLine("				THEN '03-ANU' 																")
                loConsulta.AppendLine("				ELSE '01-REG' 																")
                loConsulta.AppendLine("			END)															AS Transaccion,	")
                loConsulta.AppendLine("			(CASE WHEN Proveedores.Nacional=1 THEN 0 ELSE 1 END)   			AS Es_Importacion	")
                loConsulta.AppendLine("FROM		Ordenes_Pagos ")
                loConsulta.AppendLine("	JOIN	Proveedores ON Ordenes_Pagos.Cod_Pro = Proveedores.Cod_Pro")
                loConsulta.AppendLine("WHERE		Ordenes_Pagos.Ipos = 0")
                loConsulta.AppendLine(" 			AND Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde)
                loConsulta.AppendLine(" 			AND " & lcParametro0Hasta)
                loConsulta.AppendLine(" 			AND Ordenes_Pagos.Documento BETWEEN " & lcParametro1Desde)
                loConsulta.AppendLine(" 			AND " & lcParametro1Hasta)
                loConsulta.AppendLine(" 			AND Ordenes_Pagos.cod_pro BETWEEN " & lcParametro2Desde)
                loConsulta.AppendLine(" 			AND " & lcParametro2Hasta)
                loConsulta.AppendLine(" 			AND Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
                loConsulta.AppendLine(" 			AND " & lcParametro3Hasta)
                loConsulta.AppendLine(" 			AND Ordenes_Pagos.Status = 'Confirmado'")

                If lcParametro7Desde.ToUpper = "NO" Then
                    loConsulta.AppendLine(" 		AND Ordenes_Pagos.Mon_Imp <> 0")
                End If

                If lcParametro5Desde = "Igual" Then
                    loConsulta.AppendLine(" 		AND Ordenes_Pagos.Cod_Rev BETWEEN " & lcParametro4Desde)
                Else
                    loConsulta.AppendLine(" 		AND Ordenes_Pagos.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
                End If

                loConsulta.AppendLine(" 			AND " & lcParametro4Hasta)
                loConsulta.AppendLine("")

            End If

            '*****************	Retenciones de IVA, si aplican *************************************  
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            loConsulta.AppendLine("UPDATE		#tmpLibroCompras ")
            loConsulta.AppendLine("SET			Com_Ret = Retenciones.Com_Ret, ")
            loConsulta.AppendLine("			Fec_Ret = Retenciones.Fec_Ret, ")
            loConsulta.AppendLine("			Mon_Ret = Retenciones.Mon_Ret")
            loConsulta.AppendLine("FROM	(	SELECT		#tmpLibroCompras.Documento		AS Documento,")
            loConsulta.AppendLine("						#tmpLibroCompras.Cod_Tip		AS Cod_Tip,")
            loConsulta.AppendLine("						retenciones_documentos.num_com	AS Com_Ret,")
            loConsulta.AppendLine("						Cuentas_Pagar.fec_ini			AS Fec_Ret,")
            loConsulta.AppendLine("						Cuentas_Pagar.mon_net			AS Mon_Ret")
            loConsulta.AppendLine("			FROM		retenciones_documentos ")
            loConsulta.AppendLine("				JOIN	#tmpLibroCompras")
            loConsulta.AppendLine("					ON	#tmpLibroCompras.Documento = retenciones_documentos.doc_ori")
            loConsulta.AppendLine("					AND	#tmpLibroCompras.Codigo_Tipo = retenciones_documentos.cla_ori")
            loConsulta.AppendLine("					AND	retenciones_documentos.tip_ori = 'Cuentas_Pagar'")
            loConsulta.AppendLine("					AND	retenciones_documentos.Clase = 'IMPUESTO'")
            loConsulta.AppendLine("				JOIN	Cuentas_Pagar ")
            loConsulta.AppendLine("					ON	Cuentas_Pagar.documento = retenciones_documentos.doc_des")
            loConsulta.AppendLine("					AND	Cuentas_Pagar.cod_tip = retenciones_documentos.cla_des")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("		) AS Retenciones")
            loConsulta.AppendLine("WHERE	Retenciones.Documento = #tmpLibroCompras.documento")
            loConsulta.AppendLine("	AND	Retenciones.Cod_Tip = #tmpLibroCompras.Cod_Tip")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            '*****************	Documento de Origen de las N/CR: De devoluciones *************************************  
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE      #tmpLibroCompras")
            loConsulta.AppendLine("SET         Documento_Afectado = #tmpOrigenes.doc_ori")
            loConsulta.AppendLine("FROM    (   SELECT      Cuentas_pagar.documento,")
            loConsulta.AppendLine("                        Cuentas_pagar.cod_tip,")
            loConsulta.AppendLine("                        renglones_dproveedores.Tip_ori,")
            loConsulta.AppendLine("                        renglones_dproveedores.doc_ori")
            loConsulta.AppendLine("            FROM        renglones_dproveedores")
            loConsulta.AppendLine("                JOIN    Cuentas_pagar")
            loConsulta.AppendLine("                    ON  renglones_dproveedores.Documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("                    AND Cuentas_Pagar.cod_tip = 'N/CR' and cuentas_pagar.automatico = 1")
            loConsulta.AppendLine("                    AND renglones_dproveedores.Tip_ori  = 'Compras'")
            loConsulta.AppendLine("                JOIN    #tmpLibroCompras")
            loConsulta.AppendLine("                    ON  #tmpLibroCompras.Codigo_Tipo = 'N/CR'")
            loConsulta.AppendLine("                    AND #tmpLibroCompras.documento = Cuentas_pagar.documento")
            loConsulta.AppendLine("        ) AS #tmpOrigenes")
            loConsulta.AppendLine("WHERE   #tmpLibroCompras.Documento = #tmpOrigenes.Documento")
            loConsulta.AppendLine("    AND #tmpLibroCompras.Codigo_Tipo = #tmpOrigenes.cod_tip")
            loConsulta.AppendLine("")


            '*****************	Documento de Origen de las N/CR: De devoluciones *************************************  
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE      #tmpLibroCompras")
            loConsulta.AppendLine("SET         Documento_Afectado = #tmpOrigenes.doc_ori")
            loConsulta.AppendLine("FROM    (    SELECT     #tmpCreditos.documento,")
            loConsulta.AppendLine("                        #tmpCreditos.cod_tip,")
            loConsulta.AppendLine("                        #tmpCreditos.tip_ori,")
            loConsulta.AppendLine("                        #tmpCreditos.doc_ori")
            loConsulta.AppendLine("            FROM        cuentas_pagar AS #tmpCreditos")
            loConsulta.AppendLine("                JOIN    #tmpLibroCompras")
            loConsulta.AppendLine("                    ON  #tmpLibroCompras.Codigo_Tipo = 'N/CR'")
            loConsulta.AppendLine("                    AND #tmpLibroCompras.documento = #tmpCreditos.documento")
            loConsulta.AppendLine("            WHERE       #tmpCreditos.cod_tip = 'N/CR'")
            loConsulta.AppendLine("                    AND #tmpCreditos.automatico = 0")
            loConsulta.AppendLine("                    AND #tmpCreditos.tip_ori = 'Cuentas_Pagar'")
            loConsulta.AppendLine("                    AND #tmpCreditos.cla_ori IN ('FACT', 'N/DB')")
            loConsulta.AppendLine("        ) AS #tmpOrigenes")
            loConsulta.AppendLine("WHERE   #tmpLibroCompras.Documento = #tmpOrigenes.Documento")
            loConsulta.AppendLine("    AND #tmpLibroCompras.Codigo_Tipo = #tmpOrigenes.cod_tip")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")



            loConsulta.AppendLine("SELECT * FROM #tmpLibroCompras	")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT       Cod_Imp")
            loConsulta.AppendLine("FROM        impuestos")
            loConsulta.AppendLine("    JOIN    campos_propiedades")
            loConsulta.AppendLine("        ON  campos_propiedades.cod_reg = impuestos.cod_imp")
            loConsulta.AppendLine("        AND campos_propiedades.cod_pro = 'TASIMPRED'")
            loConsulta.AppendLine("        AND campos_propiedades.val_log = 1")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            Dim loDistribucion As System.Xml.XmlDocument
            Dim laImpuestos As System.Xml.XmlNodeList
            Dim loTabla As DataTable
            Dim loImpuestosReducidos As DataTable

            loTabla = laDatosReporte.Tables(0)
            loImpuestosReducidos = laDatosReporte.Tables(1)

            For Each loFila As DataRow In loTabla.Rows

                If Not String.IsNullOrEmpty(Trim(loFila.Item("dis_imp"))) Then

                    loDistribucion = New System.Xml.XmlDocument()
                    Try

                        loDistribucion.LoadXml(Trim(loFila.Item("dis_imp")))

                    Catch ex As Exception

                        Continue For

                    End Try

                    laImpuestos = loDistribucion.SelectNodes("impuestos/impuesto")

                    'If (loFila.Item("Cod_Tip").Equals("Orden de Pago")) Then
                    Dim lcImpuestoActual As String = ""

                    If laImpuestos.Count >= 1 Then
                        lcImpuestoActual = Trim(laImpuestos(0).SelectSingleNode("codigo").InnerText)
                        If (loImpuestosReducidos.Select("cod_imp=" & goServicios.mObtenerCampoFormatoSQL(lcImpuestoActual)).Length > 0) Then
                            loFila.Item("Impuesto1_EsReducido") = True
                        End If

                        loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
                        loFila.Item("cod_imp1") = lcImpuestoActual

                        If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                            loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText) * -1
                            loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText) * -1
                            loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText) * -1
                        Else
                            loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText)
                            loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText)
                            loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText)
                        End If

                    End If

                    If laImpuestos.Count >= 2 Then

                        lcImpuestoActual = Trim(laImpuestos(1).SelectSingleNode("codigo").InnerText)
                        If (loImpuestosReducidos.Select("cod_imp=" & goServicios.mObtenerCampoFormatoSQL(lcImpuestoActual)).Length > 0) Then
                            loFila.Item("Impuesto2_EsReducido") = True
                        End If

                        loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
                        loFila.Item("cod_imp2") = lcImpuestoActual

                        If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                            loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText) * -1
                            loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText) * -1
                            loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText) * -1
                        Else
                            loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText)
                            loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText)
                            loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText)
                        End If

                    End If

                    If laImpuestos.Count >= 3 Then

                        lcImpuestoActual = Trim(laImpuestos(2).SelectSingleNode("codigo").InnerText)
                        If (loImpuestosReducidos.Select("cod_imp=" & goServicios.mObtenerCampoFormatoSQL(lcImpuestoActual)).Length > 0) Then
                            loFila.Item("Impuesto3_EsReducido") = True
                        End If

                        loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
                        loFila.Item("cod_imp3") = Trim(laImpuestos(2).SelectSingleNode("codigo").InnerText)

                        If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                            loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText) * -1
                            loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText) * -1
                            loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText) * -1
                        Else
                            loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText)
                            loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText)
                            loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText)
                        End If

                    End If

                    'Else
                    '    If laImpuestos.Count >= 1 Then

                    '        If (loImpuestosReducidos.Select("").Length > 0) Then
                    '            loFila.Item("Impuesto1_EsReducido") = True
                    '        End If

                    '        loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
                    '        loFila.Item("cod_imp1") = Trim(laImpuestos(0).SelectSingleNode("codigo").InnerText)

                    '        If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                    '            loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText) * -1
                    '            loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText) * -1
                    '            loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText) * -1
                    '        Else
                    '            loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText)
                    '            loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText)
                    '            loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText)
                    '        End If

                    '    End If

                    '    If laImpuestos.Count >= 2 Then

                    '        If (loImpuestosReducidos.Select("").Length > 0) Then
                    '            loFila.Item("Impuesto2_EsReducido") = True
                    '        End If

                    '        loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
                    '        loFila.Item("cod_imp2") = Trim(laImpuestos(1).SelectSingleNode("codigo").InnerText)

                    '        If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                    '            loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText) * -1
                    '            loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText) * -1
                    '            loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText) * -1
                    '        Else
                    '            loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText)
                    '            loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText)
                    '            loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText)
                    '        End If

                    '    End If

                    '    If laImpuestos.Count >= 3 Then

                    '        If (loImpuestosReducidos.Select("").Length > 0) Then
                    '            loFila.Item("Impuesto3_EsReducido") = True
                    '        End If

                    '        loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
                    '        loFila.Item("cod_imp3") = Trim(laImpuestos(2).SelectSingleNode("codigo").InnerText)

                    '        If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                    '            loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText) * -1
                    '            loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText) * -1
                    '            loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText) * -1
                    '        Else
                    '            loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText)
                    '            loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText)
                    '            loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText)
                    '        End If

                    '    End If

                    'End If


                    loFila.Item("subt_imp") = CDec(loFila.Item("mon_imp3")) + CDec(loFila.Item("mon_imp2")) + CDec(loFila.Item("mon_imp1"))
                    loFila.Item("subt_exe") = CDec(loFila.Item("mon_exe1")) + CDec(loFila.Item("mon_exe2")) + CDec(loFila.Item("mon_exe3"))
                    loFila.Item("subt_bas") = CDec(loFila.Item("mon_bas1")) + CDec(loFila.Item("mon_bas2")) + CDec(loFila.Item("mon_bas3"))

                End If

            Next loFila


            loTabla.AcceptChanges()

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (loTabla.Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            Dim laDatosReporte2 As New DataSet()
            Dim loTabla2 As New DataTable()

            loTabla2 = loTabla.Copy()
            loTabla.Dispose()
            loTabla = Nothing
            laDatosReporte.Dispose()
            laDatosReporte = Nothing

            laDatosReporte2.Tables.Add(loTabla2)

            laDatosReporte2.AcceptChanges()

            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte2.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Compras_SID", laDatosReporte2)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLibro_Compras_SID.ReportSource = loObjetoReporte

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
' RJG: 05/06/13: Codigo inicial, a partir de rLibro_Compras_2013.       					'
'-------------------------------------------------------------------------------------------'
' RJG: 30/08/13: Ajuste en el título del reporte. Corrección del parámetro Teléfono del     '
'                encabezado.       					                                        '
'-------------------------------------------------------------------------------------------'
