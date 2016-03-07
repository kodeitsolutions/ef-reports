'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rLibro_Compras_2013"
'-------------------------------------------------------------------------------------------'
Partial Class rLibro_Compras_2013
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
            Dim loComandoSeleccionar As New StringBuilder()
            loComandoSeleccionar.AppendLine("DECLARE @lnCero AS DECIMAL(28, 10);")
            loComandoSeleccionar.AppendLine("SET @lnCero = CAST(0 AS DECIMAL(28, 10));")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio AS NVARCHAR(30);")
            loComandoSeleccionar.AppendLine("SET @lcVacio = N'';")

            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpLibroCompras(	Tabla	            VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Cod_Tip 	        VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Codigo_Tipo	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Documento	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Control	            VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Referencia	        VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Factura 	        VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Status		        VARCHAR(15) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Doc_Ori		        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Documento_Afectado	VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Cod_Pro		        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Fec_Ini		        DATETIME,")
            loComandoSeleccionar.AppendLine("								Tip_Doc		        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Mon_Bru 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Net 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Dis_Imp 	        VARCHAR(MAX),")
            loComandoSeleccionar.AppendLine("								Mon_Des 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Rec 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Por_Des 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Por_Rec 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Otr1 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Otr2 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Mon_Otr3 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Cod_Imp1 	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Cod_Imp2 	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Cod_Imp3 	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Com_Ret		    	VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Fec_Ret		    	DATETIME,")
            loComandoSeleccionar.AppendLine("								mon_ret 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_imp1 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_bas1 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								por_imp1 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_exe1 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_imp2 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_bas2 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								por_imp2 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_exe2 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_imp3 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_bas3 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								por_imp3 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								mon_exe3 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								subt_exe 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								subt_bas 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								subt_imp 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Exonerado 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								No_Sujeto 	        DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Sin_Derecho_CF      DECIMAL(28,10),")
            loComandoSeleccionar.AppendLine("								Nom_Pro 	        VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Rif 	            VARCHAR(20) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Tip_Pro 	        VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Transaccion		    VARCHAR(10) COLLATE DATABASE_DEFAULT)")
            loComandoSeleccionar.AppendLine("								")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpLibroCompras(	Tabla, Cod_Tip, Codigo_Tipo, Documento, Control, Referencia, Factura, Status,")
            loComandoSeleccionar.AppendLine("								Doc_Ori, Documento_Afectado, Cod_Pro, Fec_Ini, Tip_Doc, Mon_Bru, Mon_Net,")
            loComandoSeleccionar.AppendLine("								Dis_Imp, Mon_Des, Mon_Rec, Por_Des, Por_Rec, Mon_Otr1, Mon_Otr2, Mon_Otr3,")
            loComandoSeleccionar.AppendLine("								Cod_Imp1, Cod_Imp2, Cod_Imp3, mon_ret, mon_imp1, mon_bas1, por_imp1,")
            loComandoSeleccionar.AppendLine("								mon_exe1, mon_imp2, mon_bas2, por_imp2, mon_exe2, mon_imp3, mon_bas3,")
            loComandoSeleccionar.AppendLine("								por_imp3, mon_exe3, subt_exe, subt_bas, subt_imp, Exonerado, No_Sujeto,")
            loComandoSeleccionar.AppendLine("								Sin_Derecho_CF, Nom_Pro, Rif, Tip_Pro, Transaccion)")
            loComandoSeleccionar.AppendLine("SELECT		")
            loComandoSeleccionar.AppendLine("			'Compras'															AS Tabla, 		")
            loComandoSeleccionar.AppendLine("			CASE Cuentas_Pagar.Cod_Tip															")
            loComandoSeleccionar.AppendLine("			 	WHEN 'FACT' 	THEN 'Factura'													")
            loComandoSeleccionar.AppendLine("			 	WHEN 'N/CR' 	THEN 'Nota de Credito'											")
            loComandoSeleccionar.AppendLine("			 	WHEN 'N/DB' 	THEN 'Nota de Debito'											")
            loComandoSeleccionar.AppendLine("			END																	AS Cod_Tip,		")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Tip												AS Codigo_Tipo,	")
            loComandoSeleccionar.AppendLine("			CAST(Cuentas_Pagar.Documento AS CHAR(30))							AS Documento, 	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Control												AS Control, 	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Referencia											AS Referencia,	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Factura												AS Factura, 	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Status												AS Status, 		")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Doc_Ori												AS Doc_Ori, 	")
            loComandoSeleccionar.AppendLine("			@lcVacio															AS Documento_Afectado, 	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Pro												AS Cod_Pro, 	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Fec_Ini												AS Fec_Ini, 	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Tip_Doc												AS Tip_Doc, 	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Bru * -1 												")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Bru 														")
            loComandoSeleccionar.AppendLine("			END																	AS Mon_Bru,  	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net * -1 												")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Net 														")
            loComandoSeleccionar.AppendLine("			END																	AS Mon_Net,  	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Dis_Imp												AS Dis_Imp, 	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Des * -1 												")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Des 														")
            loComandoSeleccionar.AppendLine("			END																	AS Mon_Des,  	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Rec												AS Mon_Rec, 	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Por_Des												AS Por_Des, 	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Por_Rec												AS Por_Rec, 	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Otr1 * -1 												")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Otr1 													")
            loComandoSeleccionar.AppendLine("			END																	AS Mon_Otr1,  	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Otr2 * -1 												")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Otr2 													")
            loComandoSeleccionar.AppendLine("			END																	AS Mon_Otr2,  	")
            loComandoSeleccionar.AppendLine("			CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' 										")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Otr3 * -1 												")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Otr3 													")
            loComandoSeleccionar.AppendLine("			END																	AS Mon_Otr3,  	")
            loComandoSeleccionar.AppendLine("			@lcVacio															AS Cod_Imp1, 	")
            loComandoSeleccionar.AppendLine("			@lcVacio															AS Cod_Imp2, 	")
            loComandoSeleccionar.AppendLine("			@lcVacio															AS Cod_Imp3, 	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_ret, 	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_imp1, 	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_bas1, 	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS por_imp1, 	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_exe1, 	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_imp2, 	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_bas2, 	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS por_imp2, 	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS mon_exe2, 	")
            loComandoSeleccionar.AppendLine("           @lnCero 															AS mon_imp3,	")
            loComandoSeleccionar.AppendLine("           @lnCero 															AS mon_bas3,	")
            loComandoSeleccionar.AppendLine("           @lnCero 															AS por_imp3,	")
            loComandoSeleccionar.AppendLine("           @lnCero 															AS mon_exe3,	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS subt_exe,	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS subt_bas,	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS subt_imp,	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS Exonerado, 	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS No_Sujeto, 	")
            loComandoSeleccionar.AppendLine("			@lnCero 															AS Sin_Derecho_CF,	")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') 				")
            loComandoSeleccionar.AppendLine("				THEN Proveedores.Nom_Pro 														")
            loComandoSeleccionar.AppendLine("				ELSE (CASE WHEN (Cuentas_Pagar.Nom_Pro = '') 									")
            loComandoSeleccionar.AppendLine("					THEN Proveedores.Nom_Pro 													")
            loComandoSeleccionar.AppendLine("					ELSE Cuentas_Pagar.Nom_Pro 													")
            loComandoSeleccionar.AppendLine("				END) 																			")
            loComandoSeleccionar.AppendLine("			END)																AS Nom_Pro, 	")
            loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') 				")
            loComandoSeleccionar.AppendLine("				THEN Proveedores.Rif ELSE 														")
            loComandoSeleccionar.AppendLine("			    (CASE WHEN (Cuentas_Pagar.Rif = '')												")
            loComandoSeleccionar.AppendLine("					THEN Proveedores.Rif 														")
            loComandoSeleccionar.AppendLine("					ELSE Cuentas_Pagar.Rif 														")
            loComandoSeleccionar.AppendLine("			    END) 																			")
            loComandoSeleccionar.AppendLine("			 END)																AS Rif,			")
            loComandoSeleccionar.AppendLine("			Proveedores.Tip_Pro 												AS Tip_Pro,		")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Pagar.Status = 'Anulado' 										")
            loComandoSeleccionar.AppendLine("				THEN '03-ANU' 																	")
            loComandoSeleccionar.AppendLine("				ELSE '01-REG' 																	")
            loComandoSeleccionar.AppendLine("			END)																AS Transaccion	")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("	JOIN	Proveedores ON Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Documento BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.cod_pro BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Cuentas_Pagar.Status        IN ( " & lcParametro8Desde & " ) ")
            If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Pagar.Cod_Rev BETWEEN " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Pagar.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
            End If
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			AND (										")
            loComandoSeleccionar.AppendLine("						Cuentas_Pagar.Cod_Tip IN ('FACT', 'N/DB')")
            loComandoSeleccionar.AppendLine("					OR	(Cuentas_Pagar.Cod_Tip = 'N/CR' AND Cuentas_Pagar.Automatico = 1 AND Cuentas_Pagar.Tip_Ori = 'Devoluciones_Proveedores')")
            loComandoSeleccionar.AppendLine("					OR	(Cuentas_Pagar.Cod_Tip = 'N/CR' AND Cuentas_Pagar.Automatico = 0 AND Cuentas_Pagar.Cod_Rev IN ('DEVCOM', 'REBAJA') )")
            loComandoSeleccionar.AppendLine("				)										")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Pagar.cod_tip IN ('FACT', 'N/CR', 'N/DB')")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Tip      BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            Dim lcStatus As String = "Pendiente,Confirmado,Procesado,Pagado,Cerrado,Afectado,Serializado,Contabilizado,Iniciado,Conciliado,Otro,Anulado"

            If (cusAplicacion.goReportes.paParametrosIniciales(8).Equals(lcStatus)) OrElse _
             (cusAplicacion.goReportes.paParametrosIniciales(8).Equals("Anulado")) Then

                loComandoSeleccionar.AppendLine("UNION ALL")

                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine("")
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''Obtencion de las Facturas de Compra Anuladas '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                loComandoSeleccionar.AppendLine("SELECT	")
                loComandoSeleccionar.AppendLine("		'Facturas'													AS Tabla,		")
                loComandoSeleccionar.AppendLine("		'Factura'													AS cod_tip,		")
                loComandoSeleccionar.AppendLine("		@lcVacio													AS Codigo_Tipo,	")
                loComandoSeleccionar.AppendLine("		CAST(Compras.Documento AS CHAR(30))							AS Documento,	")
                loComandoSeleccionar.AppendLine("		Compras.Control												AS Control,		")
                loComandoSeleccionar.AppendLine("		@lcVacio													AS Referencia,	")
                loComandoSeleccionar.AppendLine("		Compras.Factura												AS Factura,		")
                loComandoSeleccionar.AppendLine("		Compras.Status												AS Status,		")
                loComandoSeleccionar.AppendLine("		@lcVacio													AS Doc_Ori, 	")
                loComandoSeleccionar.AppendLine("		@lcVacio													AS Documento_Afectado, 	")
                loComandoSeleccionar.AppendLine("		Compras.Cod_Pro												AS Cod_Pro, 	")
                loComandoSeleccionar.AppendLine("		Compras.Fec_Ini												AS Fec_Ini, 	")
                loComandoSeleccionar.AppendLine("		@lcVacio													AS Tip_Doc, 	")
                loComandoSeleccionar.AppendLine("		Compras.Mon_Bru												AS Mon_Bru,		")
                loComandoSeleccionar.AppendLine("		Compras.Mon_Net												AS Mon_Net,		")
                loComandoSeleccionar.AppendLine("		Compras.Dis_imp												AS Dis_imp, 	")
                loComandoSeleccionar.AppendLine("		Compras.Mon_Des1											AS Mon_Des, 	")
                loComandoSeleccionar.AppendLine("		Compras.Mon_Rec1											AS Mon_Rec, 	")
                loComandoSeleccionar.AppendLine("		Compras.Por_Des1											AS Por_Des, 	")
                loComandoSeleccionar.AppendLine("		Compras.Por_Rec1											AS Por_Rec, 	")
                loComandoSeleccionar.AppendLine("		Compras.Mon_Otr1											AS Mon_Otr1, 	")
                loComandoSeleccionar.AppendLine("		Compras.Mon_Otr2											AS Mon_Otr2, 	")
                loComandoSeleccionar.AppendLine("		Compras.Mon_Otr3											AS Mon_Otr3, 	")
                loComandoSeleccionar.AppendLine("		@lcVacio													AS Cod_Imp1, 	")
                loComandoSeleccionar.AppendLine("		@lcVacio													AS Cod_Imp2, 	")
                loComandoSeleccionar.AppendLine("		@lcVacio													AS Cod_Imp3, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS mon_ret,		")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS mon_imp1, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS mon_bas1, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS por_imp1, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS mon_exe1, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS mon_imp2, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS mon_bas2, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS por_imp2, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS mon_exe2, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS mon_imp3, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS mon_bas3, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS por_imp3, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS mon_exe3, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS subt_exe, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS subt_bas, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS subt_imp, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS Exonerado, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS No_Sujeto, 	")
                loComandoSeleccionar.AppendLine("		@lnCero 													AS Sin_Derecho_CF,	")
                loComandoSeleccionar.AppendLine("		(CASE WHEN (Proveedores.Generico = 0 AND Compras.Nom_Pro = '')				")
                loComandoSeleccionar.AppendLine("			THEN Proveedores.Nom_Pro												")
                loComandoSeleccionar.AppendLine("			ELSE (CASE WHEN (Compras.Nom_Pro = '')									")
                loComandoSeleccionar.AppendLine("				THEN Proveedores.Nom_Pro											")
                loComandoSeleccionar.AppendLine("				ELSE Compras.Nom_Pro												")
                loComandoSeleccionar.AppendLine("			END)																	")
                loComandoSeleccionar.AppendLine("		END)														AS  Nom_Pro,	")
                loComandoSeleccionar.AppendLine("		(CASE WHEN (Proveedores.Generico = 0 AND Compras.Nom_Pro = '')				")
                loComandoSeleccionar.AppendLine("			THEN Proveedores.Rif													")
                loComandoSeleccionar.AppendLine("			ELSE (CASE WHEN (Compras.Rif = '')										")
                loComandoSeleccionar.AppendLine("				THEN Proveedores.Rif												")
                loComandoSeleccionar.AppendLine("				ELSE Compras.Rif													")
                loComandoSeleccionar.AppendLine("			END)																	")
                loComandoSeleccionar.AppendLine("		END)															AS  Rif,	")
                loComandoSeleccionar.AppendLine("		Proveedores.Tip_Pro												AS Tip_Pro,	")
                loComandoSeleccionar.AppendLine("		(CASE WHEN Compras.Status = 'Anulado'										")
                loComandoSeleccionar.AppendLine("			THEN '03-ANU' 															")
                loComandoSeleccionar.AppendLine("			ELSE '01-REG' 															")
                loComandoSeleccionar.AppendLine("		END)														AS Transaccion	")
                loComandoSeleccionar.AppendLine("FROM		Compras ")
                loComandoSeleccionar.AppendLine("	JOIN	Proveedores ON Compras.Cod_Pro = Proveedores.Cod_Pro ")
                loComandoSeleccionar.AppendLine("WHERE		Compras.Fec_Ini BETWEEN " & lcParametro0Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Compras.Documento BETWEEN " & lcParametro1Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Compras.cod_pro BETWEEN " & lcParametro2Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Compras.Cod_Suc BETWEEN " & lcParametro3Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
                loComandoSeleccionar.AppendLine("			AND Compras.Status  = 'Anulado'")
                If lcParametro5Desde = "Igual" Then
                    loComandoSeleccionar.AppendLine(" 		AND Compras.Cod_Rev BETWEEN " & lcParametro4Desde)
                Else
                    loComandoSeleccionar.AppendLine(" 		AND Compras.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
                End If
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
                loComandoSeleccionar.AppendLine("")

            End If

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''Obtencion de las Ordenes de pago '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If lcParametro6Desde.ToUpper() = "SI" Then


                loComandoSeleccionar.AppendLine(" UNION ALL")
                loComandoSeleccionar.AppendLine("")
                loComandoSeleccionar.AppendLine(" SELECT	")
                loComandoSeleccionar.AppendLine("			'Orden_Pago'													AS Tabla,		")
                loComandoSeleccionar.AppendLine("			'Orden de Pago' 												AS cod_tip,		")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Codigo_Tipo,	")
                loComandoSeleccionar.AppendLine("			CAST(Ordenes_Pagos.Documento AS CHAR(30))	 					AS Documento,	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Control	 										AS Control,		")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Referencia,	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Factura	 										AS Factura,		")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Status	 										AS Status,		")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Doc_Ori, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Documento_Afectado, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Cod_Pro	 										AS Cod_Pro, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Fec_Ini	 										AS Fec_Ini, 	")
                loComandoSeleccionar.AppendLine("			'debito'														As Tip_Doc, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Bru	 										AS Mon_Bru, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Net	 										AS Mon_Net, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Dis_Imp	 										AS Dis_Imp, 	")
                loComandoSeleccionar.AppendLine("			@lnCero    				                                        AS Mon_Des, 	")
                loComandoSeleccionar.AppendLine("			@lnCero    				                                        AS Mon_Rec, 	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS Por_Des, 	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS Por_Rec, 	")
                loComandoSeleccionar.AppendLine("			@lnCero    				                                        AS Mon_Otr1, 	")
                loComandoSeleccionar.AppendLine("			@lnCero    				                                        AS Mon_Otr2, 	")
                loComandoSeleccionar.AppendLine("			@lnCero    														AS Mon_Otr3, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Cod_Imp1, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Cod_Imp2, 	")
                loComandoSeleccionar.AppendLine("			@lcVacio														AS Cod_Imp3, 	")
                loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Ret											AS Mon_Ret,		")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_imp1,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_bas1,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS por_imp1,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_exe1,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_imp2,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_bas2,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS por_imp2,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_exe2,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_imp3,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_bas3,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS por_imp3,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS mon_exe3,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS subt_exe,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS subt_bas,	")
                loComandoSeleccionar.AppendLine("			@lnCero															AS subt_imp,	")
                loComandoSeleccionar.AppendLine("			@lnCero 														AS Exonerado,	")
                loComandoSeleccionar.AppendLine("			@lnCero 														AS No_Sujeto,	")
                loComandoSeleccionar.AppendLine("			@lnCero 														AS Sin_Derecho_CF,	")
                loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '')			")
                loComandoSeleccionar.AppendLine("				THEN Proveedores.Nom_Pro													")
                loComandoSeleccionar.AppendLine("				ELSE (CASE WHEN (Ordenes_Pagos.Nom_Pro = '')								")
                loComandoSeleccionar.AppendLine("					THEN Proveedores.Nom_Pro												")
                loComandoSeleccionar.AppendLine("					ELSE Ordenes_Pagos.Nom_Pro												")
                loComandoSeleccionar.AppendLine("				END)																		")
                loComandoSeleccionar.AppendLine("			END)															AS Nom_Pro, 	")
                loComandoSeleccionar.AppendLine("			(CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '')			")
                loComandoSeleccionar.AppendLine("				THEN Proveedores.Rif														")
                loComandoSeleccionar.AppendLine("				ELSE (CASE WHEN (Ordenes_Pagos.Rif = '')									")
                loComandoSeleccionar.AppendLine("					THEN Proveedores.Rif													")
                loComandoSeleccionar.AppendLine("					ELSE Ordenes_Pagos.Rif													")
                loComandoSeleccionar.AppendLine("				END)																		")
                loComandoSeleccionar.AppendLine("			END)															AS Rif,			")
                loComandoSeleccionar.AppendLine("			Proveedores.Tip_Pro												AS Tip_Pro,		")
                loComandoSeleccionar.AppendLine("			(CASE WHEN Ordenes_Pagos.Status = 'Anulado'										")
                loComandoSeleccionar.AppendLine("				THEN '03-ANU' 																")
                loComandoSeleccionar.AppendLine("				ELSE '01-REG' 																")
                loComandoSeleccionar.AppendLine("			END)															AS Transaccion	")
                loComandoSeleccionar.AppendLine("FROM		Ordenes_Pagos ")
                loComandoSeleccionar.AppendLine("	JOIN	Proveedores ON Ordenes_Pagos.Cod_Pro = Proveedores.Cod_Pro")
                loComandoSeleccionar.AppendLine("WHERE		Ordenes_Pagos.Ipos = 0")
                loComandoSeleccionar.AppendLine(" 			AND Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Ordenes_Pagos.Documento BETWEEN " & lcParametro1Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Ordenes_Pagos.cod_pro BETWEEN " & lcParametro2Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
                loComandoSeleccionar.AppendLine(" 			AND Ordenes_Pagos.Status = 'Confirmado'")

                If lcParametro7Desde.ToUpper = "NO" Then
                    loComandoSeleccionar.AppendLine(" 		AND Ordenes_Pagos.Mon_Imp <> 0")
                End If

                If lcParametro5Desde = "Igual" Then
                    loComandoSeleccionar.AppendLine(" 		AND Ordenes_Pagos.Cod_Rev BETWEEN " & lcParametro4Desde)
                Else
                    loComandoSeleccionar.AppendLine(" 		AND Ordenes_Pagos.Cod_Rev NOT BETWEEN " & lcParametro4Desde)
                End If

                loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
                loComandoSeleccionar.AppendLine("")

            End If

            '*****************	Retenciones de IVA, si aplican *************************************  
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("UPDATE		#tmpLibroCompras ")
            loComandoSeleccionar.AppendLine("SET			Com_Ret = Retenciones.Com_Ret, ")
            loComandoSeleccionar.AppendLine("			Fec_Ret = Retenciones.Fec_Ret, ")
            loComandoSeleccionar.AppendLine("			Mon_Ret = Retenciones.Mon_Ret")
            loComandoSeleccionar.AppendLine("FROM	(	SELECT		#tmpLibroCompras.Documento		AS Documento,")
            loComandoSeleccionar.AppendLine("						#tmpLibroCompras.Cod_Tip		AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("						retenciones_documentos.num_com	AS Com_Ret,")
            loComandoSeleccionar.AppendLine("						Cuentas_Pagar.fec_ini			AS Fec_Ret,")
            loComandoSeleccionar.AppendLine("						Cuentas_Pagar.mon_net			AS Mon_Ret")
            loComandoSeleccionar.AppendLine("			FROM		retenciones_documentos ")
            loComandoSeleccionar.AppendLine("				JOIN	#tmpLibroCompras")
            loComandoSeleccionar.AppendLine("					ON	#tmpLibroCompras.Documento = retenciones_documentos.doc_ori")
            loComandoSeleccionar.AppendLine("					AND	#tmpLibroCompras.Codigo_Tipo = retenciones_documentos.cla_ori")
            loComandoSeleccionar.AppendLine("					AND	retenciones_documentos.tip_ori = 'Cuentas_Pagar'")
            loComandoSeleccionar.AppendLine("					AND	retenciones_documentos.Clase = 'IMPUESTO'")
            loComandoSeleccionar.AppendLine("				JOIN	Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("					ON	Cuentas_Pagar.documento = retenciones_documentos.doc_des")
            loComandoSeleccionar.AppendLine("					AND	Cuentas_Pagar.cod_tip = retenciones_documentos.cla_des")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("		) AS Retenciones")
            loComandoSeleccionar.AppendLine("WHERE	Retenciones.Documento = #tmpLibroCompras.documento")
            loComandoSeleccionar.AppendLine("	AND	Retenciones.Cod_Tip = #tmpLibroCompras.Cod_Tip")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")



            loComandoSeleccionar.AppendLine("SELECT * FROM #tmpLibroCompras	")
            loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento)


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            Dim loDistribucion As System.Xml.XmlDocument
            Dim laImpuestos As System.Xml.XmlNodeList
            Dim loTabla As New DataTable()

            loTabla = laDatosReporte.Tables(0)
            'If lcParametro6Desde.ToUpper() = "SI" Then

            '    If (cusAplicacion.goReportes.paParametrosIniciales(8).Equals(lcStatus)) Or _
            '     (cusAplicacion.goReportes.paParametrosIniciales(8).Equals("Anulado")) Then

            '        loTabla = laDatosReporte.Tables(0)

            '    Else
            '        loTabla = laDatosReporte.Tables(1)

            '    End If

            'Else
            '    If (cusAplicacion.goReportes.paParametrosIniciales(8).Equals(lcStatus)) Or _
            '     (cusAplicacion.goReportes.paParametrosIniciales(8).Equals("Anulado")) Then

            '        loTabla = laDatosReporte.Tables(0)

            '    Else
            '        loTabla = laDatosReporte.Tables(1)

            '    End If

            'End If

            For Each loFila As DataRow In loTabla.Rows

                If Not String.IsNullOrEmpty(Trim(loFila.Item("dis_imp"))) Then

                    loDistribucion = New System.Xml.XmlDocument()
                    Try

                        loDistribucion.LoadXml(Trim(loFila.Item("dis_imp")))

                    Catch ex As Exception

                        Continue For

                    End Try

                    laImpuestos = loDistribucion.SelectNodes("impuestos/impuesto")

                    If (loFila.Item("Cod_Tip").Equals("Orden de Pago")) Then

                        If laImpuestos.Count >= 1 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp1") = Trim(laImpuestos(0).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp1") = Trim(laImpuestos(0).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                        If laImpuestos.Count >= 2 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp2") = Trim(laImpuestos(1).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp2") = Trim(laImpuestos(1).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                        If laImpuestos.Count >= 3 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp3") = Trim(laImpuestos(2).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp3") = Trim(laImpuestos(2).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                    Else
                        If laImpuestos.Count >= 1 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp1") = Trim(laImpuestos(0).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp1") = Trim(laImpuestos(0).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp1") = CDec(laImpuestos(0).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                        If laImpuestos.Count >= 2 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp2") = Trim(laImpuestos(1).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp2") = Trim(laImpuestos(1).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp2") = CDec(laImpuestos(1).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                        If laImpuestos.Count >= 3 Then
                            If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then
                                loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp3") = Trim(laImpuestos(2).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText) * -1
                                loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText) * -1
                                loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText) * -1
                            Else
                                loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
                                loFila.Item("cod_imp3") = Trim(laImpuestos(2).SelectSingleNode("codigo").InnerText)
                                loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText)
                                loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText)
                                loFila.Item("mon_imp3") = CDec(laImpuestos(2).SelectSingleNode("monto").InnerText)
                            End If
                        End If

                    End If


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Compras_2013", laDatosReporte2)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrLibro_Compras2_Alegria.ReportSource = loObjetoReporte

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
' RJG: 15/04/13: Codigo inicial, a partir de rLibro_Compras_GrupoAlegria					'
'-------------------------------------------------------------------------------------------'
' RJG: 16/04/13: Agregado filtro para incluir solo retenciones de IVA (no ISLR ni Patente). '
'-------------------------------------------------------------------------------------------'
' RJG: 03/06/13: Se cambió la columna "Subtotal Factura" (Monto Bruto) por "Compras sin     '
'                Impuesto".                                                                 '
'-------------------------------------------------------------------------------------------'
' RJG: 30/08/13: Se ajustó el cálculo de totales para descargar los montos de documentos    '
'                anulados.                                                                  '
'-------------------------------------------------------------------------------------------'
