Imports System.Data
Partial Class rLibro_Compras2_Alegria
    Inherits vis2formularios.frmReporte
    
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


	Try	
	
			Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0),goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
			Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0),goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
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

			loComandoSeleccionar.AppendLine(" SELECT	" )
			loComandoSeleccionar.AppendLine("			'Compras' AS Tabla, " )  
			loComandoSeleccionar.AppendLine("			CASE" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'FACT' THEN 'Factura'" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'GIRO' THEN 'Giro'" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'ISRL' THEN 'ISRL'" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'N/CR' THEN 'Nota de Credito'" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'N/DB' THEN 'Nota de Debito'" )
			loComandoSeleccionar.AppendLine("			 	WHEN Cuentas_Pagar.cod_tip = 'RETIVA' THEN 'Retensión de I.V.A.'" )
			loComandoSeleccionar.AppendLine("			END AS cod_tip," )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Documento, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Control, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Referencia, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Factura, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Status, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Doc_Ori, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Pro, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Fec_Ini, " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Tip_Doc, " )
			loComandoSeleccionar.AppendLine("			CASE  " )
			loComandoSeleccionar.AppendLine("				WHEN Cuentas_Pagar.Tip_Doc = 'Credito' THEN " )
			loComandoSeleccionar.AppendLine("					Cuentas_Pagar.Mon_Bru * -1 " )
			loComandoSeleccionar.AppendLine("				ELSE  " )
			loComandoSeleccionar.AppendLine("					Cuentas_Pagar.Mon_Bru " )
			loComandoSeleccionar.AppendLine("			END AS Mon_Bru,  " )
			loComandoSeleccionar.AppendLine("			 " )
			loComandoSeleccionar.AppendLine("			CASE  " )
			loComandoSeleccionar.AppendLine("				WHEN Cuentas_Pagar.Tip_Doc = 'Credito' THEN " )
			loComandoSeleccionar.AppendLine("					Cuentas_Pagar.Mon_Net * -1 " )
			loComandoSeleccionar.AppendLine("				ELSE  " )
			loComandoSeleccionar.AppendLine("					Cuentas_Pagar.Mon_Net " )
			loComandoSeleccionar.AppendLine("			END AS Mon_Net,  " )
			loComandoSeleccionar.AppendLine("			Cuentas_Pagar.dis_imp, " )
			loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Des                                              AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Rec                                              AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Otr1                                             AS  Mon_Otr1, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Otr2                                             AS  Mon_Otr2, ")
            loComandoSeleccionar.AppendLine("           Cuentas_Pagar.Mon_Otr3                                             AS  Mon_Otr3, ")
            loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_ret, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_exe, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_bas, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_imp, " )
			loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Cuentas_Pagar.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Cuentas_Pagar.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Cuentas_Pagar.Rif = '') THEN Proveedores.Rif ELSE Cuentas_Pagar.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("			Proveedores.tip_Pro, " )
			loComandoSeleccionar.AppendLine("			(Case When Cuentas_Pagar.Status = 'Anulado' Then '03-ANU' Else '01-REG' End) as Transaccion " )
			loComandoSeleccionar.AppendLine("INTO		#tempCxP " )
			loComandoSeleccionar.AppendLine("FROM		Cuentas_Pagar, Proveedores " )
			loComandoSeleccionar.AppendLine("WHERE		Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro " )
			loComandoSeleccionar.AppendLine(" 			And Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro0Hasta )
			loComandoSeleccionar.AppendLine(" 			And Cuentas_Pagar.Documento BETWEEN " & lcParametro1Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro1Hasta )
			loComandoSeleccionar.AppendLine(" 			And Cuentas_Pagar.cod_pro BETWEEN " & lcParametro2Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 			And Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("			AND Cuentas_Pagar.Status        IN ( " & lcParametro8Desde & " ) ")
			If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Pagar.Cod_Rev between " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Cuentas_Pagar.Cod_Rev NOT between " & lcParametro4Desde)
            End If
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 			And (Cuentas_Pagar.cod_tip = 'FACT' OR Cuentas_Pagar.cod_tip = 'GIRO' OR Cuentas_Pagar.cod_tip = 'ISRL' OR ")
			loComandoSeleccionar.AppendLine(" 			Cuentas_Pagar.cod_tip = 'N/CR' OR Cuentas_Pagar.cod_tip = 'N/DB' OR Cuentas_Pagar.cod_tip = 'RETIVA')")
			loComandoSeleccionar.AppendLine("           And Cuentas_Pagar.Cod_Tip      BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            'loComandoSeleccionar.AppendLine("ORDER BY Tabla ASC,  " & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			
			Dim lcStatus AS String = "Pendiente,Confirmado,Procesado,Pagado,Cerrado,Afectado,Serializado,Contabilizado,Iniciado,Conciliado,Otro,Anulado"
			
			If (cusAplicacion.goReportes.paParametrosIniciales(8).Equals(lcStatus)) Then

					loComandoSeleccionar.AppendLine("UNION ALL")
			End If
			
			If (cusAplicacion.goReportes.paParametrosIniciales(8).Equals("Anulado")) Then

					loComandoSeleccionar.AppendLine("UNION ALL")
			End If
			
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''Obtencion de las Facturas de Compra Anuladas '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			loComandoSeleccionar.AppendLine("SELECT	" )
			loComandoSeleccionar.AppendLine("			'Facturas' AS Tabla, " )  
			loComandoSeleccionar.AppendLine("			'Factura' AS cod_tip," )
			loComandoSeleccionar.AppendLine("			Compras.Documento, " )
			loComandoSeleccionar.AppendLine("			Compras.Control, " )
			loComandoSeleccionar.AppendLine("			''								AS Referencia, " )
			loComandoSeleccionar.AppendLine("			Compras.Factura, " )
			loComandoSeleccionar.AppendLine("			Compras.Status, " )
			loComandoSeleccionar.AppendLine("			''								AS Doc_Ori, " )
			loComandoSeleccionar.AppendLine("			Compras.Cod_Pro, " )
			loComandoSeleccionar.AppendLine("			Compras.Fec_Ini, " )
			loComandoSeleccionar.AppendLine("			''								AS Tip_Doc, " )
			loComandoSeleccionar.AppendLine("			Compras.Mon_Bru," )
			loComandoSeleccionar.AppendLine("			Compras.Mon_Net," )
			loComandoSeleccionar.AppendLine("			Compras.dis_imp, " )
			loComandoSeleccionar.AppendLine("           Compras.Mon_Des1                                              AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Rec1                                              AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Otr1                                             AS  Mon_Otr1, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Otr2                                             AS  Mon_Otr2, ")
            loComandoSeleccionar.AppendLine("           Compras.Mon_Otr3                                             AS  Mon_Otr3, ")
            loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_ret, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_exe, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_bas, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_imp, " )
			loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Compras.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Compras.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Compras.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Compras.Rif = '') THEN Proveedores.Rif ELSE Compras.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("			Proveedores.tip_Pro, " )
			loComandoSeleccionar.AppendLine("			(Case When Compras.Status = 'Anulado' Then '03-ANU' Else '01-REG' End) as Transaccion " )
			loComandoSeleccionar.AppendLine("FROM		Compras, Proveedores " )
			loComandoSeleccionar.AppendLine("WHERE		Compras.Cod_Pro = Proveedores.Cod_Pro " )
			loComandoSeleccionar.AppendLine(" 			And Compras.Fec_Ini BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro0Hasta )
			loComandoSeleccionar.AppendLine(" 			And Compras.Documento BETWEEN " & lcParametro1Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro1Hasta )
			loComandoSeleccionar.AppendLine(" 			And Compras.cod_pro BETWEEN " & lcParametro2Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 			And Compras.Cod_Suc BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("			AND Compras.Status  = 'Anulado'")
			If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Compras.Cod_Rev between " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Compras.Cod_Rev NOT between " & lcParametro4Desde)
            End If
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("")
			
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''Obtencion de las Ordenes de pago '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			
			loComandoSeleccionar.AppendLine(" SELECT	" )
			loComandoSeleccionar.AppendLine("			'Orden_Pago' AS Tabla, " ) 
			loComandoSeleccionar.AppendLine("			'Orden de Pago' AS cod_tip, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Documento, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Control, " )
			loComandoSeleccionar.AppendLine("			'' AS Referencia, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Factura, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Status, " )
			loComandoSeleccionar.AppendLine("			'' AS Doc_Ori, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Cod_Pro, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Fec_Ini, " )
			loComandoSeleccionar.AppendLine("			'debito' As Tip_Doc, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Bru, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Net, " )
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.dis_imp, " )
			loComandoSeleccionar.AppendLine("           CAST(0.0 AS DECIMAL)                                            AS  Mon_Des, ")
            loComandoSeleccionar.AppendLine("           CAST(0.0 AS DECIMAL)                                            AS  Mon_Rec, ")
            loComandoSeleccionar.AppendLine("           CAST(0.0 AS DECIMAL)                                            AS  Mon_Otr1, ")
            loComandoSeleccionar.AppendLine("           CAST(0.0 AS DECIMAL)                                            AS  Mon_Otr2, ")
            loComandoSeleccionar.AppendLine("           CAST(0.0 AS DECIMAL)                                            AS  Mon_Otr3, ")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Ret, " )
            loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe1, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe2, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_imp3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_bas3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As por_imp3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As mon_exe3, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_exe, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_bas, " )
			loComandoSeleccionar.AppendLine("			CAST(0.0 AS DECIMAL) As subt_imp, " )
			loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Nom_Pro ELSE Ordenes_Pagos.Nom_Pro END) END) AS  Nom_Pro, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN (Proveedores.Generico = 0 AND Ordenes_Pagos.Nom_Pro = '') THEN Proveedores.Rif ELSE ")
            loComandoSeleccionar.AppendLine("               (CASE WHEN (Ordenes_Pagos.Rif = '') THEN Proveedores.Rif ELSE Ordenes_Pagos.Rif END) END) AS  Rif, ")
            loComandoSeleccionar.AppendLine("			Proveedores.tip_Pro, " )
			loComandoSeleccionar.AppendLine("			(Case When Ordenes_Pagos.Status = 'Anulado' Then '03-ANU' Else '01-REG' End) as Transaccion " )
			loComandoSeleccionar.AppendLine("INTO		#tempOrdenes_Pago " )
			loComandoSeleccionar.AppendLine("FROM		Ordenes_Pagos, Proveedores " )
			loComandoSeleccionar.AppendLine("WHERE		Ordenes_Pagos.Cod_Pro = Proveedores.Cod_Pro " )
			loComandoSeleccionar.AppendLine(" 			And Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro0Hasta )
			loComandoSeleccionar.AppendLine(" 			And Ordenes_Pagos.Documento BETWEEN " & lcParametro1Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro1Hasta )
			loComandoSeleccionar.AppendLine(" 			And Ordenes_Pagos.cod_pro BETWEEN " & lcParametro2Desde )
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 			And Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 			And " & lcParametro3Hasta)
			
			If lcParametro7Desde.ToUpper = "NO" Then
				loComandoSeleccionar.AppendLine(" 		And Ordenes_Pagos.Mon_Imp <> 0" )
			End If
			
			If lcParametro5Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 		AND Ordenes_Pagos.Cod_Rev between " & lcParametro4Desde)
            Else
                loComandoSeleccionar.AppendLine(" 		AND Ordenes_Pagos.Cod_Rev NOT between " & lcParametro4Desde)
            End If

            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("")
			
			
			If lcParametro6Desde.ToUpper = "SI" then 

				loComandoSeleccionar.AppendLine(" SELECT * FROM #tempCxP	")
				loComandoSeleccionar.AppendLine(" UNION ALL	")
				loComandoSeleccionar.AppendLine(" SELECT * FROM #tempOrdenes_Pago	")
				loComandoSeleccionar.AppendLine("ORDER BY Tabla ASC, " & lcOrdenamiento.Replace("Cuentas_Pagar.", " "))
								
			Else
			
				loComandoSeleccionar.AppendLine(" SELECT * FROM #tempCxP ")
			
			End If

			'me.mEscribirConsulta(loComandoSeleccionar.ToString)

	        
	        Dim loServicios As New cusDatos.goDatos
			
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")
			
			
			Dim loDistribucion	As System.Xml.XmlDocument
			Dim laImpuestos		As system.Xml.XmlNodeList
			Dim loTabla As New DataTable()
			
			If lcParametro6Desde.ToUpper = "SI" then 
			
					If (cusAplicacion.goReportes.paParametrosIniciales(8).Equals(lcStatus))	 OR	 (cusAplicacion.goReportes.paParametrosIniciales(8).Equals("Anulado")) Then
					
							loTabla = laDatosReporte.Tables(0)
							
					Else
							loTabla = laDatosReporte.Tables(1)
							
					End If

			Else
					If (cusAplicacion.goReportes.paParametrosIniciales(8).Equals(lcStatus))	 OR	 (cusAplicacion.goReportes.paParametrosIniciales(8).Equals("Anulado")) Then
					
							loTabla = laDatosReporte.Tables(0)
							
					Else
							loTabla = laDatosReporte.Tables(1)
							
					End If
			
			End If
			
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
										loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText) * -1
										loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText) * -1
										loFila.Item("mon_imp1")	= CDec(laImpuestos(0).SelectSingleNode("monto").InnerText) * -1
									Else
										loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
										loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText)
										loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText)
										loFila.Item("mon_imp1")	= CDec(laImpuestos(0).SelectSingleNode("monto").InnerText)	
									End If
								End If
									
								If laImpuestos.Count >= 2 Then 
									If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then 
										loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText) 
										loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText) * -1
										loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText) * -1
										loFila.Item("mon_imp2")	= CDec(laImpuestos(1).SelectSingleNode("monto").InnerText) * -1
									Else
										loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
										loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText)
										loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText) 
										loFila.Item("mon_imp2")	= CDec(laImpuestos(1).SelectSingleNode("monto").InnerText)
									End If
								End If
									
								If laImpuestos.Count >= 3 Then 
									If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then 
										loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText) 
										loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText) * -1
										loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText) * -1
										loFila.Item("mon_imp3")	= CDec(laImpuestos(2).SelectSingleNode("monto").InnerText) * -1
									Else
										loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
										loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText)
										loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText)
										loFila.Item("mon_imp3")	= CDec(laImpuestos(2).SelectSingleNode("monto").InnerText)
									End If
								End If
						
						Else
							   If laImpuestos.Count >= 1 Then 
									If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then 
										loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText) 
										loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText) * -1
										loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText) * -1
										loFila.Item("mon_imp1")	= CDec(laImpuestos(0).SelectSingleNode("monto").InnerText) * -1
									Else
										loFila.Item("por_imp1") = CDec(laImpuestos(0).SelectSingleNode("porcentaje").InnerText)
										loFila.Item("mon_bas1") = CDec(laImpuestos(0).SelectSingleNode("base").InnerText)
										loFila.Item("mon_exe1") = CDec(laImpuestos(0).SelectSingleNode("exento").InnerText)
										loFila.Item("mon_imp1")	= CDec(laImpuestos(0).SelectSingleNode("monto").InnerText)
									End If
								End If
									
								If laImpuestos.Count >= 2 Then 
									If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then 
										loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText) 
										loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText) * -1
										loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText) * -1
										loFila.Item("mon_imp2")	= CDec(laImpuestos(1).SelectSingleNode("monto").InnerText) * -1
									Else
										loFila.Item("por_imp2") = CDec(laImpuestos(1).SelectSingleNode("porcentaje").InnerText)
										loFila.Item("mon_bas2") = CDec(laImpuestos(1).SelectSingleNode("base").InnerText)
										loFila.Item("mon_exe2") = CDec(laImpuestos(1).SelectSingleNode("exento").InnerText) 
										loFila.Item("mon_imp2")	= CDec(laImpuestos(1).SelectSingleNode("monto").InnerText)
									End If
								End If
									
								If laImpuestos.Count >= 3 Then 
									If Trim(loFila.Item("Tip_Doc")).ToLower = "credito" Then 
										loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText) 
										loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText) * -1
										loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText) * -1
										loFila.Item("mon_imp3")	= CDec(laImpuestos(2).SelectSingleNode("monto").InnerText) * -1
									Else
										loFila.Item("por_imp3") = CDec(laImpuestos(2).SelectSingleNode("porcentaje").InnerText)
										loFila.Item("mon_bas3") = CDec(laImpuestos(2).SelectSingleNode("base").InnerText)
										loFila.Item("mon_exe3") = CDec(laImpuestos(2).SelectSingleNode("exento").InnerText) 
										loFila.Item("mon_imp3")	= CDec(laImpuestos(2).SelectSingleNode("monto").InnerText)
									End If
								End If
						
						End If
					

						loFila.Item("subt_imp") = loFila.Item("mon_imp3") + loFila.Item("mon_imp2") + loFila.Item("mon_imp1")
						loFila.Item("subt_exe") = loFila.Item("mon_exe1") + loFila.Item("mon_exe2") + loFila.Item("mon_exe3")
						loFila.Item("subt_bas") = loFila.Item("mon_bas1") + loFila.Item("mon_bas2") + loFila.Item("mon_bas3")

				End If
				
			Next lofila
 
 
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

			Dim laDatosReporte2 AS New DataSet()
			Dim loTabla2 AS New DataTable()
			
			loTabla2= loTabla.Copy()
			
			laDatosReporte2.Tables.Add(loTabla2)
			
			laDatosReporte2.AcceptChanges()
			
			
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rLibro_Compras2_Alegria", laDatosReporte2)
            
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
		
			 loObjetoReporte.Close ()

		Catch loExcepcion As Exception

		End Try
	
	End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' MAT: 03/05/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
' MAT: 12/05/11: Modificación de la Base imponible (Solo en los Totales)
'-------------------------------------------------------------------------------------------'
' MAT: 12/05/11: Ajuste para que muestre Sección de Facturas Anuladas
'-------------------------------------------------------------------------------------------'
