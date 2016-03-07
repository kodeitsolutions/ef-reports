'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPlanilla_ArcVPagos"
'-------------------------------------------------------------------------------------------'
Partial Class rPlanilla_ArcVPagos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
			
			'Parametro 1: Fecha
			'Parametro 2: Proveedor
			'Parametro 3: Concepto
			'Parametro 4: Tipo Proveedor
			'Parametro 5: Clase Proveedor
			'Parametro 6: Tipo Persona
			'Parametro 7: Solo Provedores con Retenciones
			'Parametro 8: Concepto de Proveedor
			'Parametro 9: Estatus de Proveedor
			
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
		    Dim lcParametro6Hasta As String = CStr(cusAplicacion.goReportes.paParametrosFinales(6)).Trim()
		    Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		    Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
		    
		    
		    
			Dim ldFechaInicio As Date = cusAplicacion.goReportes.paParametrosIniciales(0)
			Dim ldFechaFin As Date = cusAplicacion.goReportes.paParametrosFinales(0)
			Dim lcFechaFin AS String
			
			If	(ldFechaInicio.Year <> ldFechaFin.Year) Then
			
				  Dim ldAño As Integer = Replace(ldFechaInicio.Year,"'","")
				  lcFechaFin	= goServicios.mObtenerCampoFormatoSQL(ldAño & "1231 23:59:59.998")
			Else
					
				  lcFechaFin	=	lcParametro0Hasta
			End If

			Dim ldPrimerDia As Date = New Date(ldFechaInicio.Year, 1, 1)
			Dim lcPrimerDia As String = goServicios.mObtenerCampoFormatoSQL(ldPrimerDia, goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
         
         
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
            
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaInicio	AS DATETIME;				")
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaFin	AS DATETIME;				")
            loComandoSeleccionar.AppendLine("DECLARE @lnCero		AS DECIMAL(28, 10);			")
            loComandoSeleccionar.AppendLine("SET @ldFechaInicio	= " & lcParametro0Desde & ";	")
            loComandoSeleccionar.AppendLine("SET @ldFechaFin	= " & lcFechaFin		& "; 	")
            loComandoSeleccionar.AppendLine("SET @lnCero		= 0;							")
            loComandoSeleccionar.AppendLine("				")
            loComandoSeleccionar.AppendLine("--**********************************************************	")
            loComandoSeleccionar.AppendLine("-- Genera la tabla con el fin de cada mes						")
            loComandoSeleccionar.AppendLine("--**********************************************************	")
            loComandoSeleccionar.AppendLine("				")
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaIniMes AS DATETIME							")
            loComandoSeleccionar.AppendLine("DECLARE @ldFechaFinMes AS DATETIME							")
            loComandoSeleccionar.AppendLine("DECLARE @lnAño AS INT										")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("CREATE TABLE #tmpFechas(Dia INT, Mes INT, Año INT)																")
			loComandoSeleccionar.AppendLine("															")
			loComandoSeleccionar.AppendLine("CREATE CLUSTERED INDEX PK_tabFechas_Año_Mes_Dia ON #tmpFechas(Año, Mes, Dia)")
			loComandoSeleccionar.AppendLine("															")
            loComandoSeleccionar.AppendLine("SET @lnAño = YEAR(@ldFechaInicio);							")
            loComandoSeleccionar.AppendLine("SET @ldFechaIniMes = " & lcPrimerDia & ";					")
            loComandoSeleccionar.AppendLine("															")
            loComandoSeleccionar.AppendLine("DECLARE @lnContador INT;									")
            loComandoSeleccionar.AppendLine("SET @lnContador = 1;										")
            loComandoSeleccionar.AppendLine("															")
            loComandoSeleccionar.AppendLine("WHILE(@lnContador <= 12)									")
            loComandoSeleccionar.AppendLine("BEGIN														")
            loComandoSeleccionar.AppendLine("	SET @ldFechaIniMes = DATEADD(MONTH, 1, @ldFechaIniMes);	")
            loComandoSeleccionar.AppendLine("	SET @ldFechaFinMes = DATEADD(DAY, -1, @ldFechaIniMes);	")
            loComandoSeleccionar.AppendLine("	INSERT INTO #tmpFechas(Dia, Mes, Año) VALUES (DAY(@ldFechaFinMes), @lnContador, @lnAño);")
            loComandoSeleccionar.AppendLine("	SET @lnContador = @lnContador + 1;						")
            loComandoSeleccionar.AppendLine("END														")
            loComandoSeleccionar.AppendLine("															")
            loComandoSeleccionar.AppendLine("--**********************************************************	")
            loComandoSeleccionar.AppendLine("-- Genera un listado temporal de proveedores					")
            loComandoSeleccionar.AppendLine("--**********************************************************	")
            loComandoSeleccionar.AppendLine("				")
            loComandoSeleccionar.AppendLine("SELECT			Proveedores.Cod_Pro						AS Cod_Pro,			")
			loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro						AS Nom_Pro,			")
			loComandoSeleccionar.AppendLine("				Proveedores.Rif							AS Rif,				")
			loComandoSeleccionar.AppendLine("				Proveedores.Nit							AS Nit,				")
			loComandoSeleccionar.AppendLine("				CAST(Proveedores.Dir_Fis AS CHAR(250))	AS Dir_Fis,			")
            loComandoSeleccionar.AppendLine("				Personas.Cod_Per						AS Cod_Per,			")
            loComandoSeleccionar.AppendLine("				Personas.Nom_Per						AS Nom_Per			")
            loComandoSeleccionar.AppendLine("INTO			#tmpProveedores								")
            loComandoSeleccionar.AppendLine("FROM			Proveedores									")
            loComandoSeleccionar.AppendLine("	JOIN		Personas									")
            loComandoSeleccionar.AppendLine("			ON	Personas.Cod_Per = Proveedores.Cod_Per				")
            loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Tip BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Cla BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           AND Personas.Cod_Per BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Con BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           AND Proveedores.Status IN (" & lcParametro8Desde & ")")
	        loComandoSeleccionar.AppendLine("				")
	        loComandoSeleccionar.AppendLine("CREATE CLUSTERED INDEX PK_tmpProveedores_CodigoProveedor ON #tmpProveedores(Cod_Pro)")
	        loComandoSeleccionar.AppendLine("				")
	        
	        loComandoSeleccionar.AppendLine("")
	        loComandoSeleccionar.AppendLine("")
	        loComandoSeleccionar.AppendLine("")
	        loComandoSeleccionar.AppendLine("	        ")
	        loComandoSeleccionar.AppendLine("--**********************************************************	")
	        loComandoSeleccionar.AppendLine("-- Busca las retenciones de ISLR con sus datos.				")
	        loComandoSeleccionar.AppendLine("--**********************************************************	")
	        loComandoSeleccionar.AppendLine("         		")
	        loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori										AS Tipo_Origen,")
	        loComandoSeleccionar.AppendLine("				1															AS Dia,")
	        loComandoSeleccionar.AppendLine("				DATEPART(MONTH, Pagos.Fec_Ini)								AS Mes,")
	        loComandoSeleccionar.AppendLine("				DATEPART(YEAR, Pagos.Fec_Ini)								AS Año,")
	        loComandoSeleccionar.AppendLine("				Renglones_Pagos.Mon_Abo										AS Monto_Abonado,		")
	        loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas								AS Base_Retencion,		")
	        loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret								AS Porcentaje_Retenido,	")
	        loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Sus								AS Sustraendo_Retenido,	")
	        loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret								AS Monto_Retenido,		")
	        loComandoSeleccionar.AppendLine("				#tmpProveedores.Cod_Pro										AS Cod_Pro")
	        loComandoSeleccionar.AppendLine("INTO			#Retenciones")
	        loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
	        loComandoSeleccionar.AppendLine("		JOIN	Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori									")
	        loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento		")
	        loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento							")
			loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Cod_Ret BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("   		AND " & lcParametro2Hasta)
	        loComandoSeleccionar.AppendLine("		JOIN	Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento						")
	        loComandoSeleccionar.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori							")
	        loComandoSeleccionar.AppendLine("		LEFT JOIN Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento						")
	        loComandoSeleccionar.AppendLine("		JOIN	#tmpProveedores ON #tmpProveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro					")
	        loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'														")
	        loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'													")
	        loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'														")
	        loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
	        loComandoSeleccionar.AppendLine("         		AND " & lcFechaFin)
	        loComandoSeleccionar.AppendLine("UNION ALL																							")
	        loComandoSeleccionar.AppendLine("")
	        loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori										AS Tipo_Origen,")
	        loComandoSeleccionar.AppendLine("				1															AS Dia,")
	        loComandoSeleccionar.AppendLine("				DATEPART(MONTH, Documentos.Fec_Ini)							AS Mes,")
	        loComandoSeleccionar.AppendLine("				DATEPART(YEAR, Documentos.Fec_Ini)							AS Año,")
	        loComandoSeleccionar.AppendLine("				Documentos.Mon_Net											AS Monto_Abonado,")
	        loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas								AS Base_Retencion,")
	        loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret								AS Porcentaje_Retenido,")
	        loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Sus								AS Sustraendo_Retenido,	")
	        loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret								AS Monto_Retenido,")
	        loComandoSeleccionar.AppendLine("				#tmpProveedores.Cod_Pro										AS Cod_Pro")
	        loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
	        loComandoSeleccionar.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
	        loComandoSeleccionar.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
	        loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
			loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Cod_Ret BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("   		AND " & lcParametro2Hasta)
	        loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
	        loComandoSeleccionar.AppendLine("		JOIN	#tmpProveedores ON #tmpProveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
	        loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
	        loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
	        loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
	        loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
	        loComandoSeleccionar.AppendLine("         		AND " & lcFechaFin)
	        loComandoSeleccionar.AppendLine("")
	        
	        If (lcParametro6Hasta.ToUpper() = "SI") Then 
				loComandoSeleccionar.AppendLine("")
				loComandoSeleccionar.AppendLine("--**********************************************************	")
				loComandoSeleccionar.AppendLine("-- Descarta los proveedores sin retenciones					")
				loComandoSeleccionar.AppendLine("--**********************************************************	")
				loComandoSeleccionar.AppendLine("         		")
				loComandoSeleccionar.AppendLine("DELETE FROM #tmpProveedores")
				loComandoSeleccionar.AppendLine("WHERE NOT EXISTS (SELECT * FROM #Retenciones WHERE #Retenciones.Cod_Pro = #tmpProveedores.Cod_Pro)")
				loComandoSeleccionar.AppendLine("")
			End If
			
	        loComandoSeleccionar.AppendLine("--**********************************************************	")
	        loComandoSeleccionar.AppendLine("-- Agrupa los montos.											")
	        loComandoSeleccionar.AppendLine("--**********************************************************	")
	        loComandoSeleccionar.AppendLine("         		")
            loComandoSeleccionar.AppendLine("SELECT			--#Retenciones.Tipo_Origen									AS Tipo_Origen,			")
            loComandoSeleccionar.AppendLine("				Meses.Rif													AS Rif,					")
            loComandoSeleccionar.AppendLine("				Meses.Dia													AS Dia,")
            loComandoSeleccionar.AppendLine("				Meses.Mes													AS Mes,")
            loComandoSeleccionar.AppendLine("				Meses.Año													AS Año,")
			loComandoSeleccionar.AppendLine("				ISNULL(SUM(#Retenciones.Monto_Abonado), 0)					AS Monto_Abonado,		")
            loComandoSeleccionar.AppendLine("				ISNULL(SUM(#Retenciones.Base_Retencion), 0)				AS Base_Retencion,		")
            loComandoSeleccionar.AppendLine("				(CASE WHEN ISNULL(SUM(#Retenciones.Base_Retencion), 0)>0")
            loComandoSeleccionar.AppendLine("					THEN ROUND(SUM(#Retenciones.Monto_Retenido)*100")
            loComandoSeleccionar.AppendLine("								/SUM(#Retenciones.Base_Retencion), 2)")
            loComandoSeleccionar.AppendLine("					ELSE 0")
            loComandoSeleccionar.AppendLine("				END)														AS Porcentaje_Retenido,	")
            loComandoSeleccionar.AppendLine("				ISNULL(SUM(#Retenciones.Sustraendo_Retenido), 0)			AS Sustraendo_Retenido,	")
            loComandoSeleccionar.AppendLine("				ISNULL(SUM(#Retenciones.Monto_Retenido), 0)					AS Monto_Retenido,		")
			loComandoSeleccionar.AppendLine("				Meses.Cod_Pro												AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("				Meses.Nom_Pro												AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("				Meses.Nit													AS Nit,")
            loComandoSeleccionar.AppendLine("				Meses.Dir_Fis												AS Direccion,")
            loComandoSeleccionar.AppendLine("				Meses.Cod_Per												AS Cod_Per,")
            loComandoSeleccionar.AppendLine("				Meses.Nom_Per												AS Nom_Per,")
			loComandoSeleccionar.AppendLine("				@ldFechaInicio												AS Ini_Per_Rem,")
            loComandoSeleccionar.AppendLine("				@ldFechaFin													AS Fin_Per_Rem")
            loComandoSeleccionar.AppendLine("INTO			#TotalMensual")
            loComandoSeleccionar.AppendLine("FROM			#Retenciones")
            loComandoSeleccionar.AppendLine("	RIGHT JOIN	(	SELECT	M.Dia, M.Mes, M.Año, ")
            loComandoSeleccionar.AppendLine("							P.Cod_Pro, P.Nom_Pro, P.Rif, P.Nit,	")
            loComandoSeleccionar.AppendLine("							P.Dir_Fis, P.Cod_Per, P.Nom_Per")
            loComandoSeleccionar.AppendLine("					FROM	#tmpProveedores AS P")
			loComandoSeleccionar.AppendLine("						CROSS JOIN #tmpFechas AS M ")
			loComandoSeleccionar.AppendLine("				) AS Meses ")
            loComandoSeleccionar.AppendLine("		ON		Meses.Mes = #Retenciones.Mes")
            loComandoSeleccionar.AppendLine("		AND		Meses.Año = #Retenciones.Año")
            loComandoSeleccionar.AppendLine("		AND		Meses.Cod_Pro = #Retenciones.Cod_Pro")
			loComandoSeleccionar.AppendLine("GROUP BY		Meses.Rif, Meses.Dia, Meses.Mes, Meses.Año,")
			loComandoSeleccionar.AppendLine("				Meses.Cod_Pro, Meses.Nom_Pro, Meses.Rif, Meses.Nit,	")
			loComandoSeleccionar.AppendLine("				Meses.Dir_Fis, Meses.Cod_Per, Meses.Nom_Per")
            loComandoSeleccionar.AppendLine("")
            
            
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		T.Rif						AS Rif,					")
            loComandoSeleccionar.AppendLine("			T.Dia						AS Dia,")
            loComandoSeleccionar.AppendLine("			T.Mes						AS Mes,")
            loComandoSeleccionar.AppendLine("			T.Año						AS Año,")
            loComandoSeleccionar.AppendLine("			T.Monto_Abonado				AS Monto_Abonado,		")
            loComandoSeleccionar.AppendLine("			T.Base_Retencion			AS Base_Retencion,		")
            loComandoSeleccionar.AppendLine("			T.Porcentaje_Retenido		AS Porcentaje_Retenido,	")
            loComandoSeleccionar.AppendLine("			T.Sustraendo_Retenido		AS Sustraendo_Retenido,	")
            loComandoSeleccionar.AppendLine("			T.Monto_Retenido			AS Monto_Retenido,		")
            loComandoSeleccionar.AppendLine("			A.Retencion_Acumulada		AS Retencion_Acumulada,	")
            loComandoSeleccionar.AppendLine("			A.Impuesto_Acumulado		AS Impuesto_Acumulado,	")
            loComandoSeleccionar.AppendLine("			T.Cod_Pro					AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("			T.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("			T.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("			T.Direccion					AS Direccion,")
            loComandoSeleccionar.AppendLine("			T.Cod_Per					AS Cod_Per,")
            loComandoSeleccionar.AppendLine("			T.Nom_Per					AS Nom_Per,")
            loComandoSeleccionar.AppendLine("			T.Ini_Per_Rem				AS Ini_Per_Rem,")
            loComandoSeleccionar.AppendLine("			T.Fin_Per_Rem				AS Fin_Per_Rem")
            loComandoSeleccionar.AppendLine("FROM		#TotalMensual AS T")
            loComandoSeleccionar.AppendLine("	JOIN	(	SELECT		X.Cod_Pro, X.Mes, X.Año,")
            loComandoSeleccionar.AppendLine("							SUM(Y.Base_Retencion) AS Retencion_Acumulada,")
            loComandoSeleccionar.AppendLine("							SUM(Y.Monto_Retenido) AS Impuesto_Acumulado")
            loComandoSeleccionar.AppendLine("				FROM		#TotalMensual AS X")
            loComandoSeleccionar.AppendLine("					JOIN	#TotalMensual AS Y")
            loComandoSeleccionar.AppendLine("						ON	X.Cod_Pro = Y.Cod_Pro")
            loComandoSeleccionar.AppendLine("						AND	Y.Mes <= X.Mes")
            loComandoSeleccionar.AppendLine("				GROUP BY X.Cod_Pro, X.Mes, X.Año	")
            loComandoSeleccionar.AppendLine("			) AS A")
            loComandoSeleccionar.AppendLine("		ON	T.Cod_Pro = A.Cod_Pro ")
            loComandoSeleccionar.AppendLine("		AND T.Mes = A.Mes ")
			loComandoSeleccionar.AppendLine("ORDER BY		" & lcOrdenamiento & ", Año, Mes	")  
            loComandoSeleccionar.AppendLine("		")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpProveedores")
            loComandoSeleccionar.AppendLine("DROP TABLE #TotalMensual")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpFechas")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
            
            
            
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº 0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPlanilla_ArcVPagos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrrPlanilla_ArcVPagos.ReportSource = loObjetoReporte


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
' RJG: 19/06/12: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
