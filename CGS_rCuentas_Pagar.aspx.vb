'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rCuentas_Pagar"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rCuentas_Pagar
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFechaDesde DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaHasta DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcProDesde VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcProHasta VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcClase VARCHAR(5) = " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero DECIMAL(28, 10) = 0 ;")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio VARCHAR(10) = '';")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Órdenes de Compras'		                            AS Tipo,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Documento	                            AS Documento,")
            loComandoSeleccionar.AppendLine("		@lcVacio					                            AS Factura,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Fec_Ini		                            AS Fecha,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Cod_Pro		                            AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			                            AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Net		                            AS Monto,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Bru		                            AS Bruto,")
            loComandoSeleccionar.AppendLine("		(Ordenes_Compras.Mon_Imp1 * (Proveedores.Por_Ret/100))  AS Ret_IVA,")
            loComandoSeleccionar.AppendLine("		@lnCero						                            AS Ret_ISLR,")
            loComandoSeleccionar.AppendLine("       CASE WHEN RTRIM(Proveedores.Atributo_A) NOT LIKE '%0.0%'")
            loComandoSeleccionar.AppendLine("            THEN @lnCero")
            loComandoSeleccionar.AppendLine("            ELSE Ordenes_Compras.Mon_Bru * CONVERT(NUMERIC(18,2), Proveedores.Atributo_A)")
            loComandoSeleccionar.AppendLine("       END                                                     AS Ret_Pat,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Sal		                            AS Saldo,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Comentario	                            AS Comentario,")
            loComandoSeleccionar.AppendLine("		@lnCero						                            AS Abonado_Pago,")
            loComandoSeleccionar.AppendLine("		COALESCE((SELECT SUM(Cuentas_Pagar.Mon_Net)")
            loComandoSeleccionar.AppendLine("				 FROM Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("					JOIN Renglones_Pagos ON Cuentas_Pagar.Documento = Renglones_Pagos.Doc_Ori")
            loComandoSeleccionar.AppendLine("					JOIN Pagos ON Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("				WHERE Pagos.Ord_Com = Ordenes_Compras.Documento")
            loComandoSeleccionar.AppendLine("					AND Cuentas_Pagar.Cod_Tip = 'ADEL'), @lnCero)	AS Monto_Adel,")
            ''loComandoSeleccionar.AppendLine("					AND Cuentas_Pagar.Status <> 'Pagado'), @lnCero)	AS Monto_Adel,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Net ")
            loComandoSeleccionar.AppendLine("       - (Ordenes_Compras.Mon_Bru * (CASE WHEN RTRIM(Proveedores.Atributo_A) NOT LIKE '%0.0%'")
            loComandoSeleccionar.AppendLine("								     THEN @lnCero ELSE CONVERT(NUMERIC(18,2), Proveedores.Atributo_A) END))")
            loComandoSeleccionar.AppendLine("       - (Ordenes_Compras.Mon_Imp1 * (Proveedores.Por_Ret/100)) AS Neto_Ret,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Net ")
            loComandoSeleccionar.AppendLine("       - (Ordenes_Compras.Mon_Bru * (CASE WHEN RTRIM(Proveedores.Atributo_A) NOT LIKE '%0.0%'")
            loComandoSeleccionar.AppendLine("								      THEN @lnCero ELSE CONVERT(NUMERIC(18,2), Proveedores.Atributo_A) END))")
            loComandoSeleccionar.AppendLine("       - (Ordenes_Compras.Mon_Imp1 * (Proveedores.Por_Ret/100)) ")
            loComandoSeleccionar.AppendLine("	    - COALESCE((SELECT SUM(Cuentas_Pagar.Mon_Net)")
            loComandoSeleccionar.AppendLine("				    FROM Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("					    JOIN Renglones_Pagos ON Cuentas_Pagar.Documento = Renglones_Pagos.Doc_Ori")
            loComandoSeleccionar.AppendLine("					    JOIN Pagos ON Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("				    WHERE Pagos.Ord_Com = Ordenes_Compras.Documento")
            loComandoSeleccionar.AppendLine("					    AND Cuentas_Pagar.Cod_Tip = 'ADEL'), @lnCero)	AS Deuda,")
            ''loComandoSeleccionar.AppendLine("					    AND Cuentas_Pagar.Status <> 'Pagado'), @lnCero)	AS Monto_Adel,")
            loComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFechaDesde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFechaHasta AS DATE),103))	AS Param_Fecha,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcProDesde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro FROM Proveedores  WHERE Cod_Pro = @lcProDesde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                        AS Pro_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcProHasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro  FROM Proveedores  WHERE Cod_Pro = @lcProHasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                        AS Pro_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcClase <> 'TODOS'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Cla FROM Clases_Proveedores WHERE Cod_Cla = @lcClase)")
            loComandoSeleccionar.AppendLine("			 ELSE 'Todos'")
            loComandoSeleccionar.AppendLine("		END												        AS Clase")
            loComandoSeleccionar.AppendLine("INTO #tmpServicio")
            loComandoSeleccionar.AppendLine("FROM Ordenes_Compras")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Ordenes_Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE Ordenes_Compras.Status <> 'Pendiente'")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Documento NOT IN (SELECT Doc_Ori FROM Renglones_Compras WHERE Tip_Ori = 'ordenes_compras')")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Documento NOT IN (SELECT Doc_Ori FROM Renglones_Recepciones WHERE Tip_Ori = 'ordenes_compras')")
            ''loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Fec_Ini > '20170831'")
            ''loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Fec_Ini < @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Fec_Ini BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Cod_Pro BETWEEN @lcProDesde AND @lcProHasta")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Documento IN (SELECT Documento FROM Renglones_OCompras ")
            loComandoSeleccionar.AppendLine("										JOIN Articulos ON Renglones_OCompras.Cod_Art = Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("									  WHERE Articulos.Cod_Dep = 'SR')")
            If CStr(lcParametro2Desde).Trim() <> "'TODOS'" Then
                loComandoSeleccionar.AppendLine("   AND Proveedores.Cod_Cla = @lcClase")
            End If
            loComandoSeleccionar.AppendLine("GROUP BY Ordenes_Compras.Documento, Ordenes_Compras.Fec_Ini, Ordenes_Compras.Cod_Pro, Proveedores.Nom_Pro, Ordenes_Compras.Status,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Net, Ordenes_Compras.Mon_Sal, Ordenes_Compras.Comentario, Ordenes_Compras.Mon_Bru, Ordenes_Compras.Mon_Imp1,Proveedores.Por_Ret,Proveedores.Atributo_A")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Órdenes de Compras'		                            AS Tipo,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Documento	                            AS Documento,")
            loComandoSeleccionar.AppendLine("		@lcVacio					                            AS Factura,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Fec_Ini		                            AS Fecha,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Cod_Pro		                            AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro			                            AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Net		                            AS Monto,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Bru		                            AS Bruto,")
            loComandoSeleccionar.AppendLine("		(Ordenes_Compras.Mon_Imp1 * (Proveedores.Por_Ret/100))  AS Ret_IVA,")
            loComandoSeleccionar.AppendLine("		@lnCero						                            AS Ret_ISLR,")
            loComandoSeleccionar.AppendLine("		@lnCero						                            AS Ret_PAT,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Sal		                            AS Saldo,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Comentario	                            AS Comentario,")
            loComandoSeleccionar.AppendLine("		@lnCero						                            AS Abonado_Pago,")
            loComandoSeleccionar.AppendLine("		COALESCE((SELECT SUM(Cuentas_Pagar.Mon_Net)")
            loComandoSeleccionar.AppendLine("				 FROM Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("					JOIN Renglones_Pagos ON Cuentas_Pagar.Documento = Renglones_Pagos.Doc_Ori")
            loComandoSeleccionar.AppendLine("					JOIN Pagos ON Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("				WHERE Pagos.Ord_Com = Ordenes_Compras.Documento")
            loComandoSeleccionar.AppendLine("					AND Cuentas_Pagar.Cod_Tip = 'ADEL'), @lnCero)	AS Monto_Adel,")
            ''loComandoSeleccionar.AppendLine("					AND Cuentas_Pagar.Status <> 'Pagado'), @lnCero)	AS Monto_Adel,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Net ")
            loComandoSeleccionar.AppendLine("       - (Ordenes_Compras.Mon_Imp1 * (Proveedores.Por_Ret/100)) AS Neto_Ret,")
            loComandoSeleccionar.AppendLine("       Ordenes_Compras.Mon_Net ")
            loComandoSeleccionar.AppendLine("       - (Ordenes_Compras.Mon_Imp1 * (Proveedores.Por_Ret/100))")
            loComandoSeleccionar.AppendLine("	    - COALESCE((SELECT SUM(Cuentas_Pagar.Mon_Net)")
            loComandoSeleccionar.AppendLine("				    FROM Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("					    JOIN Renglones_Pagos ON Cuentas_Pagar.Documento = Renglones_Pagos.Doc_Ori")
            loComandoSeleccionar.AppendLine("					    JOIN Pagos ON Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("				    WHERE Pagos.Ord_Com = Ordenes_Compras.Documento")
            loComandoSeleccionar.AppendLine("					    AND Cuentas_Pagar.Cod_Tip = 'ADEL'), @lnCero)	AS Deuda,")
            ''loComandoSeleccionar.AppendLine("					    AND Cuentas_Pagar.Status <> 'Pagado'), @lnCero)	AS Monto_Adel,")
            loComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFechaDesde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFechaHasta AS DATE),103))	AS Param_Fecha,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcProDesde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro FROM Proveedores  WHERE Cod_Pro = @lcProDesde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                        AS Pro_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcProHasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro  FROM Proveedores  WHERE Cod_Pro = @lcProHasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                        AS Pro_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcClase <> 'TODOS'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Cla FROM Clases_Proveedores WHERE Cod_Cla = @lcClase)")
            loComandoSeleccionar.AppendLine("			 ELSE 'Todos'")
            loComandoSeleccionar.AppendLine("		END												        AS Clase")
            loComandoSeleccionar.AppendLine("FROM Ordenes_Compras")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Ordenes_Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE (Ordenes_Compras.Status <> 'Pendiente'")
            loComandoSeleccionar.AppendLine("	OR (Ordenes_Compras.Documento IN (SELECT Doc_Ori FROM Renglones_Compras WHERE Tip_Ori = 'ordenes_compras')")
            loComandoSeleccionar.AppendLine("       AND Ordenes_Compras.Status = 'Afectado'))")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Documento NOT IN (SELECT Doc_Ori FROM Renglones_Recepciones WHERE Tip_Ori = 'ordenes_compras')")
            ''loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Fec_Ini > '20170831'")
            ''loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Fec_Ini < @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Fec_Ini BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Ordenes_Compras.Cod_Pro BETWEEN @lcProDesde AND @lcProHasta")
            loComandoSeleccionar.AppendLine("   AND Ordenes_Compras.Documento NOT IN (SELECT Documento FROM #tmpServicio)")
            If CStr(lcParametro2Desde).Trim() <> "'TODOS'" Then
                loComandoSeleccionar.AppendLine("   AND Proveedores.Cod_Cla = @lcClase")
            End If
            loComandoSeleccionar.AppendLine("GROUP BY Ordenes_Compras.Documento, Ordenes_Compras.Fec_Ini, Ordenes_Compras.Cod_Pro, Proveedores.Nom_Pro, Ordenes_Compras.Status,")
            loComandoSeleccionar.AppendLine("		Ordenes_Compras.Mon_Net, Ordenes_Compras.Mon_Sal, Ordenes_Compras.Comentario, Ordenes_Compras.Mon_Bru, Ordenes_Compras.Mon_Imp1,Proveedores.Por_Ret,Proveedores.Atributo_A")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	'Facturas'										AS Tipo,")
            loComandoSeleccionar.AppendLine("		Compras.Documento								AS Documento,")
            loComandoSeleccionar.AppendLine("		Compras.Factura									AS Factura,")
            loComandoSeleccionar.AppendLine("		Compras.Fec_Ini									AS Fecha,")
            loComandoSeleccionar.AppendLine("		Compras.Cod_Pro									AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro								AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Compras.Mon_Net									AS Monto,")
            loComandoSeleccionar.AppendLine("		Compras.Mon_Bru									AS Bruto,")
            loComandoSeleccionar.AppendLine("		COALESCE(RETIVA.Mon_Sal,@lnCero)				AS Ret_IVA,")
            loComandoSeleccionar.AppendLine("		COALESCE(ISLR.Mon_Sal,@lnCero)					AS Ret_ISLR,")
            loComandoSeleccionar.AppendLine("		COALESCE(RETPAT.Mon_Sal,@lnCero)				AS Ret_PAT,")
            loComandoSeleccionar.AppendLine("		Compras.Mon_Sal									AS Saldo,")
            loComandoSeleccionar.AppendLine("		Compras.Comentario								AS Comentario,")
            loComandoSeleccionar.AppendLine("		COALESCE(SUM(Renglones_Pagos.Mon_Abo), @lnCero)	AS Abonado_Pago,")
            loComandoSeleccionar.AppendLine("		@lnCero											AS Monto_Adel,")
            loComandoSeleccionar.AppendLine("		Compras.Mon_Net ")
            loComandoSeleccionar.AppendLine("		- (COALESCE(RETIVA.Mon_Sal,@lnCero)+")
            loComandoSeleccionar.AppendLine("			COALESCE(ISLR.Mon_Sal,@lnCero)+")
            loComandoSeleccionar.AppendLine("			COALESCE(RETPAT.Mon_Sal,@lnCero))		    AS Neto_Ret,")
            loComandoSeleccionar.AppendLine("		Compras.Mon_Net ")
            loComandoSeleccionar.AppendLine("		- (COALESCE(RETIVA.Mon_Sal,@lnCero)+")
            loComandoSeleccionar.AppendLine("			COALESCE(ISLR.Mon_Sal,@lnCero)+")
            loComandoSeleccionar.AppendLine("			COALESCE(RETPAT.Mon_Sal,@lnCero))")
            loComandoSeleccionar.AppendLine("		- COALESCE(SUM(Renglones_Pagos.Mon_Abo), @lnCero)	AS Deuda,")
            loComandoSeleccionar.AppendLine("		CONCAT(CONVERT(VARCHAR(12),CAST(@ldFechaDesde AS DATE),103), ' - ',  CONVERT(VARCHAR(12),CAST(@ldFechaHasta AS DATE),103))	AS Param_Fecha,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcProDesde <> ''")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro FROM Proveedores  WHERE Cod_Pro = @lcProDesde)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                AS Pro_Desde,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcProHasta <> 'zzzzzzz'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Pro  FROM Proveedores  WHERE Cod_Pro = @lcProHasta)")
            loComandoSeleccionar.AppendLine("			 ELSE '' END				                AS Pro_Hasta,")
            loComandoSeleccionar.AppendLine("		CASE WHEN @lcClase <> 'TODOS'")
            loComandoSeleccionar.AppendLine("			 THEN (SELECT Nom_Cla FROM Clases_Proveedores WHERE Cod_Cla = @lcClase)")
            loComandoSeleccionar.AppendLine("			 ELSE 'Todos'")
            loComandoSeleccionar.AppendLine("		END												AS Clase")
            loComandoSeleccionar.AppendLine("FROM Compras")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Pagar AS RETIVA ON RETIVA.Doc_Ori = Compras.Documento")
            loComandoSeleccionar.AppendLine("		AND RETIVA.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("		AND RETIVA.Tip_Ori = 'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("		AND RETIVA.Cod_Pro = Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND RETIVA.Status NOT IN ('Anulado', 'Pagado')")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Pagar AS ISLR ON ISLR.Doc_Ori = Compras.Documento")
            loComandoSeleccionar.AppendLine("		AND ISLR.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("		AND ISLR.Tip_Ori = 'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("		AND ISLR.Cod_Pro = Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND ISLR.Status NOT IN ('Anulado', 'Pagado')")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Cuentas_Pagar AS RETPAT ON RETPAT.Doc_Ori = Compras.Documento")
            loComandoSeleccionar.AppendLine("		AND RETPAT.Cod_Tip = 'RETPAT'")
            loComandoSeleccionar.AppendLine("		AND RETPAT.Tip_Ori = 'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("		AND RETPAT.Cod_Pro = Compras.Cod_Pro")
            loComandoSeleccionar.AppendLine("		AND RETPAT.Status NOT IN ('Anulado', 'Pagado')")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Renglones_Pagos ")
            loComandoSeleccionar.AppendLine("		INNER JOIN Pagos ")
            loComandoSeleccionar.AppendLine("			ON (Renglones_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("			AND Pagos.Status = 'Confirmado')	")
            loComandoSeleccionar.AppendLine("	ON Renglones_Pagos.Doc_Ori = Compras.Documento ")
            loComandoSeleccionar.AppendLine("		AND Renglones_Pagos.Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("WHERE Compras.Mon_Sal > 0	")
            'loComandoSeleccionar.AppendLine("	AND Compras.Fec_Ini < @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Compras.Fec_Ini BETWEEN @ldFechaDesde AND @ldFechaHasta")
            loComandoSeleccionar.AppendLine("	AND Compras.Cod_Pro BETWEEN @lcProDesde AND @lcProHasta")
            If CStr(lcParametro2Desde).Trim() <> "'TODOS'" Then
                loComandoSeleccionar.AppendLine("   AND Proveedores.Cod_Cla = @lcClase")
            End If
            loComandoSeleccionar.AppendLine("GROUP BY Compras.Documento, Compras.Factura, Compras.Control, Compras.Fec_Ini, Compras.Cod_Pro, Proveedores.Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Compras.Status, Compras.Mon_Net	, Compras.Mon_Sal, Compras.Comentario, RETIVA.Mon_Sal, ISLR.Mon_Sal,")
            loComandoSeleccionar.AppendLine("		RETPAT.Mon_Sal, Compras.Mon_Bru")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT * FROM #tmpServicio")
            loComandoSeleccionar.AppendLine("ORDER BY Tipo, Cod_Pro, Documento")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpServicio")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rCuentas_Pagar", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rCuentas_Pagar.ReportSource = loObjetoReporte

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
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' MJP: 16/07/08: Codigo inicial
'-------------------------------------------------------------------------------------------'
' GS:  14/03/16: Cambio a Listado de Artículos.
'-------------------------------------------------------------------------------------------'

