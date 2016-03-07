'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPlanilla_ArcVOrdenesPagos"
'-------------------------------------------------------------------------------------------'
Partial Class rPlanilla_ArcVOrdenesPagos
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
		    Dim lcParametro6Hasta As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
		    Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
		    
		    
		    
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
            loComandoSeleccionar.AppendLine("DECLARE @tabFechas AS TABLE(Dia INT, Mes INT, Año INT)		")
			loComandoSeleccionar.AppendLine("															")
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
            loComandoSeleccionar.AppendLine("	INSERT INTO @tabFechas(Dia, Mes, Año) VALUES (DAY(@ldFechaFinMes), @lnContador, @lnAño);")
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
			loComandoSeleccionar.AppendLine("				CAST(Proveedores.Dir_Fis AS CHAR(250))	AS Direccion,		")
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
	        loComandoSeleccionar.AppendLine("				")
	        loComandoSeleccionar.AppendLine("CREATE CLUSTERED INDEX PK_tmpProveedores_CodigoProveedor ON #tmpProveedores(Cod_Pro)")
	        loComandoSeleccionar.AppendLine("				")
	        
	        
            loComandoSeleccionar.AppendLine("SELECT		Ordenes_Pagos.Documento					AS Documento,	")
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Cod_Pro					AS Cod_Pro,		")
			loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Mon_Net					AS Mon_Net		")
            loComandoSeleccionar.AppendLine("INTO		#tmpConceptosGastos")
            loComandoSeleccionar.AppendLine("FROM		Ordenes_Pagos")
            loComandoSeleccionar.AppendLine("JOIN		Renglones_OPagos ")
            loComandoSeleccionar.AppendLine("		ON	Renglones_OPagos.Documento =  Ordenes_Pagos.Documento	")
			loComandoSeleccionar.AppendLine("		AND Renglones_OPagos.Cod_Con BETWEEN " & lcParametro7Desde)
			loComandoSeleccionar.AppendLine("       AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Ordenes_Pagos.Documento,Cod_Pro,Ordenes_Pagos.Mon_Net	")
            
            
            
            loComandoSeleccionar.AppendLine("SELECT			MONTH(Ordenes_Pagos.Fec_Ini)			AS Mes,		")
            loComandoSeleccionar.AppendLine("				YEAR(Ordenes_Pagos.Fec_Ini)				AS Año,		")
            
            If (cusAplicacion.goReportes.paParametrosFinales(6)="Si") Then
			   loComandoSeleccionar.AppendLine("			ISNULL(SUM(")
			   loComandoSeleccionar.AppendLine("				CASE WHEN Retenciones_Documentos.Renglon=1 ")
			   loComandoSeleccionar.AppendLine("					THEN #tmpConceptosGastos.mon_net	")
			   loComandoSeleccionar.AppendLine("					ELSE 0		")
			   loComandoSeleccionar.AppendLine("				END), 0)								AS Monto_Abonado,")
			Else
				loComandoSeleccionar.AppendLine("			SUM(CASE WHEN ISNULL(Retenciones_Documentos.Renglon,1) = 1")
			   	loComandoSeleccionar.AppendLine("					THEN Ordenes_Pagos.mon_net ")
			   	loComandoSeleccionar.AppendLine("					ELSE 0		")
			   	loComandoSeleccionar.AppendLine("				END)									AS Monto_Abonado,")
			End If
            
            loComandoSeleccionar.AppendLine("				SUM(Retenciones_Documentos.Mon_Bas)		AS Base_Retencion,		")
            loComandoSeleccionar.AppendLine("				SUM(Retenciones_Documentos.Mon_Sus)		AS Sustraendo_Retenido,	")
            loComandoSeleccionar.AppendLine("				SUM(Retenciones_Documentos.Mon_Ret)		AS Monto_Retenido,		")
            loComandoSeleccionar.AppendLine("				#tmpProveedores.Cod_Pro					AS Cod_Pro				")
            loComandoSeleccionar.AppendLine("INTO			#tmpTotales														")
            loComandoSeleccionar.AppendLine("FROM			Ordenes_Pagos													")
			loComandoSeleccionar.AppendLine("		JOIN	#tmpConceptosGastos ")
			loComandoSeleccionar.AppendLine("			ON	#tmpConceptosGastos.Documento = Ordenes_Pagos.Documento	")

            If (cusAplicacion.goReportes.paParametrosFinales(6)="Si") Then
					loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos											")
					loComandoSeleccionar.AppendLine("			ON	#tmpConceptosGastos.Documento = Retenciones_Documentos.documento")
					loComandoSeleccionar.AppendLine("			AND	Retenciones_Documentos.Tip_Ori	= 'Ordenes_Pagos'				")
					loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.clase	= 'ISLR'						")
					loComandoSeleccionar.AppendLine("		JOIN	#tmpProveedores													")
					loComandoSeleccionar.AppendLine("			ON	#tmpConceptosGastos.Cod_Pro = #tmpProveedores.Cod_Pro			")
					loComandoSeleccionar.AppendLine("		JOIN	Retenciones														")
					loComandoSeleccionar.AppendLine("			ON	Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret			")
					loComandoSeleccionar.AppendLine("			AND Retenciones.Cod_Ret BETWEEN " & lcParametro2Desde)
					loComandoSeleccionar.AppendLine("   	    AND " & lcParametro2Hasta)
					
            Else
				    loComandoSeleccionar.AppendLine("		LEFT JOIN	Retenciones_Documentos											")
					loComandoSeleccionar.AppendLine("				ON	#tmpConceptosGastos.Documento = Retenciones_Documentos.documento")
					loComandoSeleccionar.AppendLine("				AND	Retenciones_Documentos.Tip_Ori	= 'Ordenes_Pagos'				")
					loComandoSeleccionar.AppendLine("				AND Retenciones_Documentos.clase	= 'ISLR'						")
					loComandoSeleccionar.AppendLine("		JOIN		#tmpProveedores													")
					loComandoSeleccionar.AppendLine("				ON	#tmpConceptosGastos.Cod_Pro = #tmpProveedores.Cod_Pro			")
					loComandoSeleccionar.AppendLine("		FULL JOIN	Retenciones														")
					loComandoSeleccionar.AppendLine("				ON	Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret			")
					loComandoSeleccionar.AppendLine("				AND Retenciones.Cod_Ret BETWEEN " & lcParametro2Desde)
					loComandoSeleccionar.AppendLine("       		AND " & lcParametro2Hasta)
					
            		
            End If
            loComandoSeleccionar.AppendLine("WHERE			Ordenes_Pagos.Status = 'Confirmado'								")
			loComandoSeleccionar.AppendLine("           	AND Ordenes_Pagos.Fec_Ini 	BETWEEN @ldFechaInicio AND @ldFechaFin	")
			loComandoSeleccionar.AppendLine("GROUP BY		MONTH(Ordenes_Pagos.Fec_Ini), YEAR(Ordenes_Pagos.Fec_Ini), 		")
			loComandoSeleccionar.AppendLine("				#tmpProveedores.Cod_Pro			")
			loComandoSeleccionar.AppendLine("ORDER BY		#tmpProveedores.Cod_Pro, MONTH(Ordenes_Pagos.Fec_Ini), YEAR(Ordenes_Pagos.Fec_Ini) ASC")
            loComandoSeleccionar.AppendLine("				")
            loComandoSeleccionar.AppendLine("				")
            
            
            
            loComandoSeleccionar.AppendLine("SELECT			Meses_Proveedor.Dia											AS Dia,					")
            loComandoSeleccionar.AppendLine("				Meses_Proveedor.Mes											AS Mes,					")
            loComandoSeleccionar.AppendLine("				Meses_Proveedor.Año											AS Año,					")
            loComandoSeleccionar.AppendLine("				Meses_Proveedor.Cod_Pro										AS Cod_Pro,				")
            loComandoSeleccionar.AppendLine("				Meses_Proveedor.Nom_Pro										AS Nom_Pro,				")
            loComandoSeleccionar.AppendLine("				Meses_Proveedor.Rif											AS Rif,					")
            loComandoSeleccionar.AppendLine("				Meses_Proveedor.Nit											AS Nit,					")
            loComandoSeleccionar.AppendLine("				Meses_Proveedor.Direccion	 								AS Direccion, 			")
            loComandoSeleccionar.AppendLine("				Meses_Proveedor.Cod_Per 									AS Cod_Per, 			")
            loComandoSeleccionar.AppendLine("				Meses_Proveedor.Nom_Per										AS Nom_Per,				")
            loComandoSeleccionar.AppendLine("				ISNULL(#tmpTotales.Monto_Abonado, @lnCero)					AS Monto_Abonado,		")
            loComandoSeleccionar.AppendLine("				ISNULL(#tmpTotales.Base_Retencion, @lnCero)					AS Base_Retencion,		")
            loComandoSeleccionar.AppendLine("				ISNULL(#tmpTotales.Sustraendo_Retenido, @lnCero)			AS Sustraendo_Retenido,	")
            loComandoSeleccionar.AppendLine("				CASE WHEN (ISNULL(#tmpTotales.Base_Retencion, @lnCero)>0)")
            loComandoSeleccionar.AppendLine("					THEN ROUND(ISNULL(#tmpTotales.Monto_Retenido, @lnCero)*100/ISNULL(#tmpTotales.Base_Retencion, @lnCero), 2)")
            loComandoSeleccionar.AppendLine("					ELSE @lnCero")
            loComandoSeleccionar.AppendLine("				END															AS Porcentaje_Retenido, ")
            loComandoSeleccionar.AppendLine("				ISNULL(#tmpTotales.Monto_Retenido, @lnCero)					AS Monto_Retenido		")
            loComandoSeleccionar.AppendLine("INTO			#tmpFinal")
            loComandoSeleccionar.AppendLine("FROM			#tmpTotales												")
            loComandoSeleccionar.AppendLine("	RIGHT JOIN (														")
            loComandoSeleccionar.AppendLine("					SELECT		#tmpProveedores.Cod_Pro, 				")
            loComandoSeleccionar.AppendLine("								#tmpProveedores.Nom_Pro, 				")
            loComandoSeleccionar.AppendLine("								#tmpProveedores.Rif, 					")
            loComandoSeleccionar.AppendLine("								#tmpProveedores.Nit, 					")
            loComandoSeleccionar.AppendLine("								#tmpProveedores.Direccion, 				")
            loComandoSeleccionar.AppendLine("								#tmpProveedores.Cod_Per, 				")
            loComandoSeleccionar.AppendLine("								#tmpProveedores.Nom_Per, 				")
            loComandoSeleccionar.AppendLine("								Meses_Años.Mes, 						")
            loComandoSeleccionar.AppendLine("								Meses_Años.Año, 						")
            loComandoSeleccionar.AppendLine("								Meses_Años.Dia							")
            loComandoSeleccionar.AppendLine("					FROM		#tmpProveedores							")
            loComandoSeleccionar.AppendLine("						CROSS JOIN @tabFechas AS Meses_Años				")
            loComandoSeleccionar.AppendLine("				) AS Meses_Proveedor						")
            loComandoSeleccionar.AppendLine("		ON Meses_Proveedor.Cod_Pro = #tmpTotales.Cod_Pro	")
            loComandoSeleccionar.AppendLine("		AND Meses_Proveedor.Mes = #tmpTotales.Mes			")
            loComandoSeleccionar.AppendLine("		AND Meses_Proveedor.Año = #tmpTotales.Año			")
            loComandoSeleccionar.AppendLine("ORDER BY		Cod_Pro, Año, Mes							")
            loComandoSeleccionar.AppendLine("				")
            loComandoSeleccionar.AppendLine("				")
            
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT		Final.Mes,")
            loComandoSeleccionar.AppendLine("			Final.Cod_Pro,	")
            loComandoSeleccionar.AppendLine("			SUM(Acumulado.Base_Retencion)AS Retencion_Acumulada,")
            loComandoSeleccionar.AppendLine("			SUM(Acumulado.Monto_Retenido)AS Impuesto_Acumulado	")
            loComandoSeleccionar.AppendLine("INTO		#tmpTotales_Acumulados								")
            loComandoSeleccionar.AppendLine("FROM		#tmpFinal AS Final									")
            loComandoSeleccionar.AppendLine("	JOIN	#tmpFinal AS Acumulado								")
            loComandoSeleccionar.AppendLine("		ON	Acumulado.Mes <= Final.Mes ")
            loComandoSeleccionar.AppendLine("		AND Acumulado.Cod_Pro = Final.Cod_Pro")
            loComandoSeleccionar.AppendLine("GROUP BY	Final.Mes, Final.Cod_Pro")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
			
			loComandoSeleccionar.AppendLine("SELECT		#tmpFinal.*,")
			loComandoSeleccionar.AppendLine("			#tmpTotales_Acumulados.Retencion_Acumulada,")
			loComandoSeleccionar.AppendLine("			#tmpTotales_Acumulados.Impuesto_Acumulado,")
			loComandoSeleccionar.AppendLine("			CONVERT(datetime," & lcParametro0Desde & ",103) AS Ini_Per_Rem,")
			loComandoSeleccionar.AppendLine("			CONVERT(datetime," & lcFechaFin & ",103) AS Fin_Per_Rem")
			loComandoSeleccionar.AppendLine("FROM 		#tmpFinal")
			loComandoSeleccionar.AppendLine("	JOIN 	#tmpTotales_Acumulados")
			loComandoSeleccionar.AppendLine("		ON	(#tmpTotales_Acumulados.Mes = #tmpFinal.Mes)")
			loComandoSeleccionar.AppendLine("		AND	(#tmpTotales_Acumulados.Cod_Pro = #tmpFinal.Cod_Pro)")
			loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento & ", Año, Mes	")
			loComandoSeleccionar.AppendLine("				")
            
            


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
            
           'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPlanilla_ArcVOrdenesPagos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrrPlanilla_ArcVOrdenesPagos.ReportSource = loObjetoReporte


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
' CMS: 21/05/09: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' CMS: 28/07/09: Se modificó la consulta de modo que se obtuvieron por separado los			'
'				 proveedores y los beneficiarios y luego se unieron los resultados.			'
'				 Verificacion de registros.													'
'				 Metodo de Ordenamiento														'
'-------------------------------------------------------------------------------------------'
' CMS: 29/07/09: Se Renonbre de Relación Global de ISLR Relativo a Relación Global de ISLR	'
'				 Retenido																	'
'-------------------------------------------------------------------------------------------'
' RJG: 20/03/10: Agregado el filtro para que distinga retenciones de IVA de las de ISLR.	'
'-------------------------------------------------------------------------------------------'
' MAT: 09/06/11: Programación del rpt del reporte.	Ajuste del Select						'
'-------------------------------------------------------------------------------------------'
' MAT: 12/06/11: Agregado Filtro Con/Sin Retenciones (Ajuste del Select)					'
'-------------------------------------------------------------------------------------------'
' MAT: 27/06/11: Agregado Filtro Concepto de Gastos (Ajuste del Select)						'
'-------------------------------------------------------------------------------------------'
' RJG: 18/06/12: Se ajustó el SELECT para corregir cálculo del acumulado.					'
'-------------------------------------------------------------------------------------------'
' RJG: 18/06/12: Se cambió el nombre "rPlanilla_ArcOrdenesPagos" por						'
'				 "rPlanilla_ArcVOrdenesPagos".												'
'-------------------------------------------------------------------------------------------'
