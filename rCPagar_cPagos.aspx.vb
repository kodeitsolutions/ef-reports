'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCPagar_cPagos"
'-------------------------------------------------------------------------------------------'
Partial Class rCPagar_cPagos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpCuentasPagar(	Documento CHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                               Cod_Tip CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                               Fec_Ini DATETIME, ")
            loComandoSeleccionar.AppendLine("                               Fec_Fin DATETIME, ")
            loComandoSeleccionar.AppendLine("								Cod_Pro CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("								Nom_Pro CHAR(100) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("								Cod_Ven CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("								Cod_Tra CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("								Cod_Mon CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("								Control CHAR(30) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("								Tip_Doc CHAR(30) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("								Tip_Ori CHAR(50) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("								Mon_Bru DECIMAL(28, 10), ")
            loComandoSeleccionar.AppendLine("								Mon_Imp DECIMAL(28, 10), ")
            loComandoSeleccionar.AppendLine("								Mon_Net DECIMAL(28, 10), ")
            loComandoSeleccionar.AppendLine("								Mon_Sal DECIMAL(28, 10), ")
            loComandoSeleccionar.AppendLine("								Comentario VARCHAR(MAX) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("								Est_Pag CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("								Num_Pag CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("								Fec_Pag DATETIME, ")
            loComandoSeleccionar.AppendLine("								Mon_Pag DECIMAL(28, 10), ")
            loComandoSeleccionar.AppendLine("								Mon_Abo DECIMAL(28, 10),")
            loComandoSeleccionar.AppendLine("								Des_Pag DECIMAL(28, 10))")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpCuentasPagar(	Documento, Cod_Tip, Fec_Ini, Fec_Fin, Cod_Pro, Nom_Pro, ")
            loComandoSeleccionar.AppendLine("								Cod_Ven, Cod_Tra, Cod_Mon, Control, Tip_Doc, Tip_Ori,")
            loComandoSeleccionar.AppendLine("								Mon_Bru, Mon_Imp, Mon_Net, Mon_Sal, Comentario, ")
            loComandoSeleccionar.AppendLine("								Est_Pag, Num_Pag, Fec_Pag, Mon_Pag, Mon_Abo, Des_Pag)")
            loComandoSeleccionar.AppendLine("SELECT		Cuentas_Pagar.Documento							AS Documento, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Tip							AS Cod_Tip, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Fec_Ini							AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Fec_Fin							AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Pro							AS Cod_Pro, ")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro								AS Nom_Pro, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Ven							AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Tra							AS Cod_Tra, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Mon							AS Cod_Mon, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Control							AS Control, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Tip_Doc							AS Tip_Doc, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Tip_Ori							AS Tip_Ori,")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Bru *(-1) ")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Bru ")
            loComandoSeleccionar.AppendLine("			END)											AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Imp1							AS Mon_Imp, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Net *(-1) ")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Net ")
            loComandoSeleccionar.AppendLine("			END)											AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' ")
            loComandoSeleccionar.AppendLine("				THEN Cuentas_Pagar.Mon_Sal *(-1) ")
            loComandoSeleccionar.AppendLine("				ELSE Cuentas_Pagar.Mon_Sal ")
            loComandoSeleccionar.AppendLine("			END)											AS Mon_Sal,  ")
            loComandoSeleccionar.AppendLine("			CASE    ")
            loComandoSeleccionar.AppendLine("				WHEN (DATALENGTH(Cuentas_Pagar.Comentario) > 1) AND (DATALENGTH(Cuentas_Pagar.Notas) > 1) ")
            loComandoSeleccionar.AppendLine("					THEN '- '+CAST(Cuentas_Pagar.Comentario AS  VARCHAR(1000))+CHAR(13)+'- '+CAST(Cuentas_Pagar.Notas AS  VARCHAR(1000)) ")
            loComandoSeleccionar.AppendLine("				WHEN (DATALENGTH(Cuentas_Pagar.Comentario) > 1) AND (DATALENGTH(Cuentas_Pagar.Notas) <= 1) ")
            loComandoSeleccionar.AppendLine("					THEN '- '+CAST(Cuentas_Pagar.Comentario AS  VARCHAR(1000))   ")
            loComandoSeleccionar.AppendLine("				WHEN (DATALENGTH(Cuentas_Pagar.Comentario) <= 1) AND (DATALENGTH(Cuentas_Pagar.Notas) > 1) ")
            loComandoSeleccionar.AppendLine("					THEN '- '+CAST(Cuentas_Pagar.Notas AS  VARCHAR(1000))   ")
            loComandoSeleccionar.AppendLine("				ELSE ''    ")
            loComandoSeleccionar.AppendLine("			END												AS Comentario,")
            loComandoSeleccionar.AppendLine("			Pagos.Status									AS Est_Pag, ")
            loComandoSeleccionar.AppendLine("			Pagos.Documento									AS Num_Pag, ")
            loComandoSeleccionar.AppendLine("			Renglones_Pagos.Registro						AS Fec_Pag, ")
            loComandoSeleccionar.AppendLine("			Renglones_Pagos.Mon_Net							AS Mon_Pag, ")
            loComandoSeleccionar.AppendLine("			Renglones_Pagos.Mon_Abo							AS Mon_Abo, ")
            loComandoSeleccionar.AppendLine("			Pagos.Mon_Des									AS Des_Pag ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("	JOIN	Proveedores ")
            loComandoSeleccionar.AppendLine("		ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro ")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Pagos ")
            loComandoSeleccionar.AppendLine("		ON	Cuentas_Pagar.Documento = Renglones_Pagos.Doc_Ori ")
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Cod_Tip	= Renglones_Pagos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("	JOIN	Pagos ")
            loComandoSeleccionar.AppendLine("		ON	Pagos.Documento = Renglones_Pagos.Documento  ")
            loComandoSeleccionar.AppendLine("WHERE		(Cuentas_Pagar.Cod_Tip = 'ADEL' OR Cuentas_Pagar.Tip_Ori <> 'Pagos')")
            loComandoSeleccionar.AppendLine("	AND NOT	(Cuentas_Pagar.Cod_Tip IN ('ISLR', 'RETIVA', 'RETPAT') AND  Cuentas_Pagar.Tip_Ori IN ('Pagos'))")
			loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Documento		BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Fec_Ini		BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Tip		BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Pro		BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Ven		BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Zon			BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Status		IN (" & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Tip			BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Cla			BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Tra		BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Mon		BETWEEN " & lcParametro10Desde & " AND " & lcParametro10Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Suc		BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta & "")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Rev		BETWEEN " & lcParametro12Desde & " AND " & lcParametro12Hasta & "")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	#tmpCuentasPagar.Documento, #tmpCuentasPagar.Cod_Tip, #tmpCuentasPagar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("		#tmpCuentasPagar.Fec_Fin, #tmpCuentasPagar.Cod_Pro, #tmpCuentasPagar.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("		#tmpCuentasPagar.Cod_Ven, #tmpCuentasPagar.Cod_Tra, #tmpCuentasPagar.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("		#tmpCuentasPagar.Control, #tmpCuentasPagar.Tip_Doc, #tmpCuentasPagar.Tip_Ori,")
            loComandoSeleccionar.AppendLine("		#tmpCuentasPagar.Mon_Bru, #tmpCuentasPagar.Mon_Imp, #tmpCuentasPagar.Mon_Net, ")
            loComandoSeleccionar.AppendLine("		#tmpCuentasPagar.Mon_Sal, #tmpCuentasPagar.Comentario, ")
            loComandoSeleccionar.AppendLine("		#tmpCuentasPagar.Est_Pag, #tmpCuentasPagar.Num_Pag, #tmpCuentasPagar.Fec_Pag, ")
            loComandoSeleccionar.AppendLine("		#tmpCuentasPagar.Mon_Pag, #tmpCuentasPagar.Mon_Abo, #tmpCuentasPagar.Des_Pag AS Mon_Des,")
            loComandoSeleccionar.AppendLine("		COALESCE(Retenciones.Mon_Ret_IVA, 0)		AS Mon_Ret_IVA, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Retenciones.Mon_Ret_ISLR, 0)		AS Mon_Ret_ISLR, ")
            loComandoSeleccionar.AppendLine("		COALESCE(Retenciones.Mon_Ret_PATENTE, 0)	AS Mon_Ret_PATENTE,")
            loComandoSeleccionar.AppendLine("		CAST(0 AS DECIMAL)			AS	Neto, ")
            loComandoSeleccionar.AppendLine("		CAST(0 AS DECIMAL)			AS	Saldo ")
            loComandoSeleccionar.AppendLine("FROM	#tmpCuentasPagar")
            loComandoSeleccionar.AppendLine("	LEFT JOIN(	SELECT	Retenciones_Documentos.Doc_Ori							AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("						Retenciones_Documentos.Cod_Tip							AS Cod_Tip,")
            loComandoSeleccionar.AppendLine("						SUM(CASE ")
            loComandoSeleccionar.AppendLine("							WHEN (Retenciones_Documentos.Clase = 'IMPUESTO')")
            loComandoSeleccionar.AppendLine("							THEN Retenciones_Documentos.Mon_Ret")
            loComandoSeleccionar.AppendLine("							ELSE 0")
            loComandoSeleccionar.AppendLine("						END)													AS Mon_Ret_IVA, ")
            loComandoSeleccionar.AppendLine("						SUM(CASE ")
            loComandoSeleccionar.AppendLine("							WHEN (Retenciones_Documentos.Clase = 'ISLR')")
            loComandoSeleccionar.AppendLine("							THEN Retenciones_Documentos.Mon_Ret")
            loComandoSeleccionar.AppendLine("							ELSE 0")
            loComandoSeleccionar.AppendLine("						END)													AS Mon_Ret_ISLR, ")
            loComandoSeleccionar.AppendLine("						SUM(CASE ")
            loComandoSeleccionar.AppendLine("							WHEN (Retenciones_Documentos.Clase = 'PATENTE')")
            loComandoSeleccionar.AppendLine("							THEN Retenciones_Documentos.Mon_Ret")
            loComandoSeleccionar.AppendLine("							ELSE 0")
            loComandoSeleccionar.AppendLine("						END)													AS Mon_Ret_PATENTE ")
            loComandoSeleccionar.AppendLine("				FROM	Retenciones_Documentos ")
            loComandoSeleccionar.AppendLine("					JOIN #tmpCuentasPagar AS CxP_Retenida")
            loComandoSeleccionar.AppendLine("					ON	Retenciones_Documentos.Doc_Ori = CxP_Retenida.Documento	  ")
            loComandoSeleccionar.AppendLine("					AND	Retenciones_Documentos.Cod_Tip = CxP_Retenida.Cod_Tip	 ")
            loComandoSeleccionar.AppendLine("					AND Retenciones_Documentos.Origen = 'Pagos'")
            loComandoSeleccionar.AppendLine("				WHERE	Retenciones_Documentos.Clase IN ('ISLR', 'IMPUESTO', 'PATENTE')")
            loComandoSeleccionar.AppendLine("				GROUP BY Retenciones_Documentos.Doc_Ori, Retenciones_Documentos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("		) AS Retenciones")
            loComandoSeleccionar.AppendLine("	ON	Retenciones.Doc_Ori = #tmpCuentasPagar.Documento	  ")
            loComandoSeleccionar.AppendLine("	AND	Retenciones.Cod_Tip = #tmpCuentasPagar.Cod_Tip	 ")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Pro ASC, Cod_Tip ASC, Documento ASC, Num_Pag ASC")
            'loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento & "")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpCuentasPagar")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            
   '         loComandoSeleccionar.AppendLine("SELECT")
   '         loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Documento	AS Doc_Ret, ")
   '         loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Doc_Ori	AS Doc_Ori, ")
   '         loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Tip_Ori	AS Tip_Ori, ")
   '       	loComandoSeleccionar.AppendLine(" 		SUM(CASE	WHEN Retenciones_Documentos.Clase = 'IMPUESTO' THEN Retenciones_Documentos.Mon_Ret ELSE  0.00 END) AS  Mon_Ret_IVA, ")
			'loComandoSeleccionar.AppendLine(" 		SUM(CASE	WHEN Retenciones_Documentos.Clase = 'ISLR' THEN Retenciones_Documentos.Mon_Ret ELSE  0.00 END) AS  Mon_Ret_ISLR, ")
			'loComandoSeleccionar.AppendLine(" 		SUM(CASE	WHEN Retenciones_Documentos.Clase = 'PATENTE' THEN Retenciones_Documentos.Mon_Ret ELSE  0.00 END) AS  Mon_Ret_PATENTE ")
			'loComandoSeleccionar.AppendLine("INTO	#tablaPagos1")
			'loComandoSeleccionar.AppendLine("FROM   Cuentas_Pagar,Renglones_Pagos,Retenciones_Documentos,Proveedores")
			'loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Documento	BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta & "")
			'loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Documento	=	Retenciones_Documentos.Doc_Ori")
			'loComandoSeleccionar.AppendLine("		AND	Renglones_Pagos.Documento	=	Retenciones_Documentos.Documento")
			'loComandoSeleccionar.AppendLine("		AND	Renglones_Pagos.Renglon	=	Retenciones_Documentos.Renglon")
			'loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Cod_Tip	=	Retenciones_Documentos.Cod_Tip")
			'loComandoSeleccionar.AppendLine("		AND	Retenciones_Documentos.Clase IN ('IMPUESTO','ISLR','PATENTE')")
			'loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Cod_Pro		=	Proveedores.Cod_Pro ")
			'loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Documento		BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Fec_Ini		BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Tip		BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Pro		BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Ven		BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Zon			BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Status		IN (" & lcParametro6Desde & ")")
   '         loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Tip			BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Cla			BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Tra		BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Mon		BETWEEN " & lcParametro10Desde & " AND " & lcParametro10Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Suc		BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Rev		BETWEEN " & lcParametro12Desde & " AND " & lcParametro12Hasta & "")
			'loComandoSeleccionar.AppendLine(" GROUP BY Retenciones_Documentos.Documento, Retenciones_Documentos.Doc_Ori,Retenciones_Documentos.Tip_Ori") 
			'loComandoSeleccionar.AppendLine("")
			
   '         loComandoSeleccionar.AppendLine("SELECT")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Documento, ")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Tip, ")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini, ")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Fin, ")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Pro, ")
   '         loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro, ")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Ven, ")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Tra, ")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Mon, ")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Control, ")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Tip_Doc, ")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Tip_Ori,")
   '         loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Bru *(-1) ELSE Cuentas_Pagar.Mon_Bru END) AS Mon_Bru, ")
   '         loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Mon_Imp1, ")
   '         loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Net *(-1) ELSE Cuentas_Pagar.Mon_Net END) AS Mon_Net, ")
   '         loComandoSeleccionar.AppendLine("		(CASE WHEN Cuentas_Pagar.Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Sal *(-1) ELSE Cuentas_Pagar.Mon_Sal END) AS Mon_Sal,  ")
   '         loComandoSeleccionar.AppendLine("		CASE    ")
   '         loComandoSeleccionar.AppendLine("			WHEN (DATALENGTH(Cuentas_Pagar.Comentario) > 1) AND (DATALENGTH(Cuentas_Pagar.Notas) > 1) THEN '- '+CAST(Cuentas_Pagar.Comentario AS  VARCHAR(1000))+CHAR(13)+'- '+CAST(Cuentas_Pagar.Notas AS  VARCHAR(1000)) ")
   '         loComandoSeleccionar.AppendLine("			WHEN (DATALENGTH(Cuentas_Pagar.Comentario) > 1) AND (DATALENGTH(Cuentas_Pagar.Notas) <= 1) THEN '- '+CAST(Cuentas_Pagar.Comentario AS  VARCHAR(1000))   ")
   '         loComandoSeleccionar.AppendLine("			WHEN (DATALENGTH(Cuentas_Pagar.Comentario) <= 1) AND (DATALENGTH(Cuentas_Pagar.Notas) > 1) THEN '- '+CAST(Cuentas_Pagar.Notas AS  VARCHAR(1000))   ")
   '         loComandoSeleccionar.AppendLine("			ELSE ''    ")
   '         loComandoSeleccionar.AppendLine("		END AS Comentario, ")
   '         loComandoSeleccionar.AppendLine("		Renglones_Pagos.Documento	AS	Num_Pag, ")
   '         loComandoSeleccionar.AppendLine("		Renglones_Pagos.Registro		AS	Fec_Pag, ")
   '         loComandoSeleccionar.AppendLine("		Renglones_Pagos.Mon_Net		AS	Mon_Pag, ")            
   '         loComandoSeleccionar.AppendLine("		Renglones_Pagos.Mon_Abo		AS	Mon_Abo, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos1.Doc_Ret		AS	Doc_Ret, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos1.Doc_Ori		AS	Origen, ")
			'loComandoSeleccionar.AppendLine("		ISNULL (#tablaPagos1.Mon_Ret_IVA,0) AS Mon_Ret_IVA, ")
			'loComandoSeleccionar.AppendLine("		ISNULL (#tablaPagos1.Mon_Ret_ISLR,0) AS Mon_Ret_ISLR, ")
			'loComandoSeleccionar.AppendLine("		ISNULL (#tablaPagos1.Mon_Ret_PATENTE,0) AS Mon_Ret_PATENTE, ")
   '         loComandoSeleccionar.AppendLine("		ISNULL (Descuentos_Documentos.Mon_Des,0)	AS	Mon_Des ")
   '         loComandoSeleccionar.AppendLine("INTO	#tablaPagos")
   '         loComandoSeleccionar.AppendLine("FROM   Proveedores, Cuentas_Pagar")
   '         loComandoSeleccionar.AppendLine("LEFT JOIN Renglones_Pagos ON (Cuentas_Pagar.Documento	=	Renglones_Pagos.Doc_Ori AND Cuentas_Pagar.Cod_Tip	=	Renglones_Pagos.Cod_Tip) ")
   '         loComandoSeleccionar.AppendLine("LEFT JOIN #tablaPagos1 ON (Renglones_Pagos.Documento	=	#tablaPagos1.Doc_Ret AND Cuentas_Pagar.Documento	=	#tablaPagos1.Doc_Ori AND #tablaPagos1.Tip_Ori  =  'Cuentas_Pagar') ")
			'loComandoSeleccionar.AppendLine("LEFT JOIN Descuentos_Documentos ON (Renglones_Pagos.Documento	=	Descuentos_Documentos.Documento AND Renglones_Pagos.Renglon	=	Descuentos_Documentos.Renglon AND Cuentas_Pagar.Documento	=	Descuentos_Documentos.Doc_Ori AND Descuentos_Documentos.Tip_Ori  =  'Cuentas_Pagar') ")
			'loComandoSeleccionar.AppendLine("JOIN Pagos ON (Renglones_Pagos.Documento = Pagos.Documento  AND Pagos.Automatico = 0) ")    
			'loComandoSeleccionar.AppendLine("WHERE	Cuentas_Pagar.Cod_Pro			=	Proveedores.Cod_Pro ")
   '         loComandoSeleccionar.AppendLine("		AND (Cuentas_Pagar.cod_tip = 'ADEL' OR Cuentas_Pagar.Tip_Ori <> 'Pagos')")
   '         loComandoSeleccionar.AppendLine("		AND NOT(Cuentas_Pagar.Cod_Tip IN ('ISLR', 'RETIVA', 'RETPAT') AND  Cuentas_Pagar.Tip_Ori IN ('Pagos'))")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Documento		BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Fec_Ini		BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Tip		BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Pro		BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Ven		BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Zon			BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Status		IN (" & lcParametro6Desde & ")")
   '         loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Tip			BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Cla			BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Tra		BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Mon		BETWEEN " & lcParametro10Desde & " AND " & lcParametro10Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Suc		BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta & "")
   '         loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Rev		BETWEEN " & lcParametro12Desde & " AND " & lcParametro12Hasta & "")
   '         loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento & "")
   '         loComandoSeleccionar.AppendLine("")
   '         loComandoSeleccionar.AppendLine("SELECT")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Documento		AS	Documento, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Cod_Tip			AS	Cod_Tip, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Fec_Ini			AS	Fec_Ini, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Fec_Fin			AS	Fec_Fin, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Cod_Pro			AS	Cod_Pro, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Nom_Pro			AS	Nom_Pro, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Cod_Ven			AS	Cod_Ven, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Cod_Tra			AS	Cod_Tra, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Cod_Mon			AS	Cod_Mon, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Control			AS	Control, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Tip_Doc			AS	Tip_Doc, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Mon_Bru			AS	Mon_Bru, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Mon_Imp1		AS	Mon_Imp, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Mon_Net			AS	Mon_Net, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Mon_Sal			AS	Mon_Sal,  ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Tip_Ori			AS  Tip_Ori,  ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Comentario		AS	Comentario, ")
   '         loComandoSeleccionar.AppendLine("		Pagos.Status				AS	Est_Pag, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Num_Pag			AS	Num_Pag, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Fec_Pag			AS	Fec_Pag, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Mon_Abo			AS	Mon_Abo, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Mon_Des			AS	Mon_Des, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Mon_Ret_IVA		AS	Mon_Ret_IVA, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Mon_Ret_ISLR	AS	Mon_Ret_ISLR, ")
   '         loComandoSeleccionar.AppendLine("		#tablaPagos.Mon_Ret_PATENTE		AS	Mon_Ret_PATENTE, ")
   '         loComandoSeleccionar.AppendLine("		CAST(0 AS DECIMAL)			AS	Neto, ")
   '         loComandoSeleccionar.AppendLine("		CAST(0 AS DECIMAL)				AS	Saldo, ")            
   '         loComandoSeleccionar.AppendLine("		(CASE WHEN Pagos.Status IN ('Confirmado','Afectado') THEN #tablaPagos.Mon_Net	ELSE 0.00 END)	AS	Mon_Pag")
   '         loComandoSeleccionar.AppendLine("FROM	#tablaPagos,Pagos")
   '         loComandoSeleccionar.AppendLine("WHERE	#tablaPagos.Num_Pag	=	Pagos.Documento ") 
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos
            
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
            
            If laDatosReporte.Tables(0).Rows.Count > 0   Then
					
					Dim Tabla As DataTable  = laDatosReporte.Tables(0)
					Dim Filas As Integer = Tabla.Rows.Count
					Dim Documento_Anterior As String = Tabla.Rows(0).Item("Documento")
					
					Tabla.Rows(0).Item("Neto")	=	Tabla.Rows(0).Item("Mon_Net")
					Tabla.Rows(0).Item("Saldo")	=	Tabla.Rows(0).Item("Mon_Sal")
					
					For i As Integer = 1 To Filas-1 
					
						If	(Documento_Anterior = Tabla.Rows(i).Item("Documento"))
					
							Tabla.Rows(i).Item("Neto")	=	0
							Tabla.Rows(i).Item("Saldo")	=	0
						Else
							Tabla.Rows(i).Item("Neto")	=	Tabla.Rows(i).Item("Mon_Net")
							Tabla.Rows(i).Item("Saldo")	=	Tabla.Rows(i).Item("Mon_Sal")
						End If
						
						Documento_Anterior  = Tabla.Rows(i).Item("Documento")
						
					Next i
					
			End If
            
           'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCPagar_cPagos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCPagar_cPagos.ReportSource = loObjetoReporte

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
' MAT: 05/01/11: Programacion inicial 
'-------------------------------------------------------------------------------------------'
' RJG: 28/08/13: Programacion inicial 
'-------------------------------------------------------------------------------------------'
