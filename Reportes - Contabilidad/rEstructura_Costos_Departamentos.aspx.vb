'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEstructura_Costos_Departamentos"
'-------------------------------------------------------------------------------------------'
Partial Class rEstructura_Costos_Departamentos
    Inherits vis2Formularios.frmReporte

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
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden


            Dim loComandoSeleccionar As New StringBuilder()
            loComandoSeleccionar.AppendLine("DECLARE @lcCadena NVARCHAR(MAX);         ")
            loComandoSeleccionar.AppendLine("DECLARE @lcAux NVARCHAR(MAX);         ")
            loComandoSeleccionar.AppendLine("DECLARE @lcAux2 NVARCHAR(MAX);         ")
            loComandoSeleccionar.AppendLine("DECLARE @lcAux3 NVARCHAR(MAX);         ")
            loComandoSeleccionar.AppendLine("DECLARE @lcAux4 NVARCHAR(MAX);         ")
            loComandoSeleccionar.AppendLine("DECLARE @lnTotal Decimal(28,10);         ")
            loComandoSeleccionar.AppendLine("DECLARE @lnEntero INT;         ")
            loComandoSeleccionar.AppendLine("DECLARE @lnEntero1 INT;         ")
            loComandoSeleccionar.AppendLine("DECLARE @llBandera BIT;         ")

            loComandoSeleccionar.AppendLine("SELECT         ")
            loComandoSeleccionar.AppendLine("                       Departamentos.Cod_Dep															AS Cod_Dep,")
            loComandoSeleccionar.AppendLine("                       Departamentos.Nom_Dep															AS Nom_Dep,")
            loComandoSeleccionar.AppendLine("                       SUM(COALESCE(Renglones_Facturas.Can_Art1,0))									AS Cantidad_Ventas,")
            loComandoSeleccionar.AppendLine("                       SUM(COALESCE(Renglones_Facturas.Mon_Net,0))										AS Ventas_Totales,")
            loComandoSeleccionar.AppendLine("                       SUM(COALESCE((Renglones_Facturas.Cos_Ult1 * Renglones_Facturas.Can_Art1),0))	AS Costo_Venta,")
            loComandoSeleccionar.AppendLine("						SUM (SUM(COALESCE(Renglones_Facturas.Mon_Net,0))) OVER()						AS Tot_Net,")
            loComandoSeleccionar.AppendLine("						(Case SUM (SUM(COALESCE(Renglones_Facturas.Mon_Net,0))) OVER()")
            loComandoSeleccionar.AppendLine("							WHEN 0 THEN 0")
            loComandoSeleccionar.AppendLine("							ELSE (SUM(COALESCE(Renglones_Facturas.Mon_Net,0))")
            loComandoSeleccionar.AppendLine("									/SUM (SUM(COALESCE(Renglones_Facturas.Mon_Net,0))) OVER())*100")
            loComandoSeleccionar.AppendLine("						END)																			AS  Por_Ven")
            loComandoSeleccionar.AppendLine("INTO #TmpPrincipal")
            loComandoSeleccionar.AppendLine("FROM                   Facturas")
            loComandoSeleccionar.AppendLine("       JOIN            Renglones_Facturas")
            loComandoSeleccionar.AppendLine("            	    ON	Renglones_Facturas.Documento = Facturas.Documento")
            loComandoSeleccionar.AppendLine("		            AND Facturas.Status IN ('Confirmado','Afectado','Procesado')")
            loComandoSeleccionar.AppendLine("			        AND Renglones_Facturas.Cod_Art	BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("				    AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			        AND Facturas.Fec_Ini			BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("				    AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			        AND Facturas.Cod_Mon			BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("				    AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("			        AND Facturas.Cod_Rev			BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    			    AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("			        AND Facturas.Cod_Suc			BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    			    AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("       JOIN            Articulos")
            loComandoSeleccionar.AppendLine("            	    ON	Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine("	 		        AND Articulos.Cod_Sec			BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("				    AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			        AND Articulos.Cod_Cla			BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("				    AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("			        AND Articulos.Cod_Tip			BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("				    AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       RIGHT JOIN		Departamentos ON (Departamentos.Cod_Dep  = Articulos.Cod_Dep)")
            loComandoSeleccionar.AppendLine("			        AND Departamentos.Cod_Dep		BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("				    AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	            Departamentos.Cod_Dep, Departamentos.Nom_Dep")
            loComandoSeleccionar.AppendLine("ORDER BY	            " & lcOrdenamiento)

            loComandoSeleccionar.AppendLine("SELECT SUM((Renglones_Comprobantes.mon_deb - Renglones_Comprobantes.mon_hab))			AS [HAB],         ")
            loComandoSeleccionar.AppendLine("		Cuentas_Gastos.Cod_Gas,          ")
            loComandoSeleccionar.AppendLine("		Cuentas_Gastos.nom_Gas,          ")
            loComandoSeleccionar.AppendLine("		SUM(SUM((Renglones_Comprobantes.mon_deb - Renglones_Comprobantes.mon_hab))) OVER() Total         ")
            loComandoSeleccionar.AppendLine("INTO #tmpGastos         ")
            loComandoSeleccionar.AppendLine("FROM comprobantes          ")
            loComandoSeleccionar.AppendLine("	JOIN		Renglones_Comprobantes ON Renglones_Comprobantes.documento = Comprobantes.Documento         ")
            loComandoSeleccionar.AppendLine("		AND		Renglones_Comprobantes.Adicional = Comprobantes.Adicional         ")
            loComandoSeleccionar.AppendLine("		AND     Comprobantes.Status IN ('Confirmado','Pendiente')")
            loComandoSeleccionar.AppendLine("		AND     Renglones_Comprobantes.Fec_Ori	BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("		AND     " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND     Comprobantes.Cod_Mon			BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("		AND     " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("		AND     Comprobantes.Cod_Rev			BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("    	AND     " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("		AND     Comprobantes.Cod_Suc			BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("    	AND     " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("		JOIN	Cuentas_Contables ON Cuentas_Contables.Cod_cue = Renglones_Comprobantes.Cod_cue         ")
            loComandoSeleccionar.AppendLine("		JOIN	Cuentas_Gastos ON Cuentas_Gastos.cod_gas = Cuentas_Contables.cod_gas         ")
            loComandoSeleccionar.AppendLine("	GROUP BY Cuentas_Gastos.Cod_Gas, Cuentas_Gastos.nom_Gas")


            loComandoSeleccionar.AppendLine("SELECT @lcAux=  COALESCE(@lcAux + '],[', '[') + RTRIM(nom_Gas)         ")
            loComandoSeleccionar.AppendLine("				FROM #tmpGastos          ")
            loComandoSeleccionar.AppendLine("					WHERE RTRIM(nom_Gas) <> ' '         ")
            loComandoSeleccionar.AppendLine("				ORDER BY nom_Gas         ")

            loComandoSeleccionar.AppendLine("SET @lcAux = @lcAux + ']'         ")

            loComandoSeleccionar.AppendLine("IF @lcAux is NULL         ")
            loComandoSeleccionar.AppendLine("	SET @lcAux = '[ ]'         ")
            loComandoSeleccionar.AppendLine("SET @lcCadena = N'	SELECT *          ")
            loComandoSeleccionar.AppendLine("					INTO #tmpGastosTotales          ")
            loComandoSeleccionar.AppendLine("					FROM  #tmpGastos         ")
            loComandoSeleccionar.AppendLine("						PIVOT (SUM(Hab) FOR [nom_Gas] IN ('+@lcAux+') ) AS Tabla'         ")

            loComandoSeleccionar.AppendLine("SET @llBandera = 1;         ")
            loComandoSeleccionar.AppendLine("SET @lnEntero = 2;         ")
            loComandoSeleccionar.AppendLine("SET @lnEntero1 = 1;         ")
            loComandoSeleccionar.AppendLine("SET @lcAux2 ='';         ")
            loComandoSeleccionar.AppendLine("SET @lcAux3 ='';         ")
            loComandoSeleccionar.AppendLine("SET @lcAux4 ='';         ")


            loComandoSeleccionar.AppendLine("WHILE (@llBandera = 1 and @lcAux <> '[ ]' and @lnEntero1<= 5)         ")
            loComandoSeleccionar.AppendLine("BEGIN         ")
            loComandoSeleccionar.AppendLine("	IF SUBSTRING(@lcAux,@lnEntero,1) = ']'         ")
            loComandoSeleccionar.AppendLine("            BEGIN         ")
            loComandoSeleccionar.AppendLine("		SET @lcAux3 = SUBSTRING(@lcAux,1,@lnEntero)         ")
            loComandoSeleccionar.AppendLine("		SET @lcAux2 = @lcAux2 + ', ' + @lcAux3+'*Por_Ven/100 AS '+'Columna'+CAST(@lnEntero1 AS Varchar(20))         ")
            loComandoSeleccionar.AppendLine("		SET @lcAux4 = @lcAux4 + ', '+ ''''+ SUBSTRING(@lcAux3,2,LEN(@lcAux3)-2)+ ''' AS Nombre' +CAST(@lnEntero1 AS Varchar(20))         ")
            loComandoSeleccionar.AppendLine("		SET @lnEntero1 = @lnEntero1 + 1         ")
            loComandoSeleccionar.AppendLine("		IF @lnEntero = LEN(@lcAux) SET @llBandera = 0;         ")
            loComandoSeleccionar.AppendLine("		ELSE SET @lcAux = SUBSTRING(@lcAux,@lnEntero+2,LEN(@lcAux))         ")
            loComandoSeleccionar.AppendLine("		SET @lnEntero = 0         ")
            loComandoSeleccionar.AppendLine("            End         ")
            loComandoSeleccionar.AppendLine("	SET @lnEntero = @lnEntero + 1         ")
            loComandoSeleccionar.AppendLine("END;         ")

            loComandoSeleccionar.AppendLine("WHILE @lnEntero1 <= 5")
            loComandoSeleccionar.AppendLine("BEGIN")
            loComandoSeleccionar.AppendLine("		SET @lcAux4 = @lcAux4 + ', '+ '0.00 AS Columna' +CAST(@lnEntero1 AS Varchar(20))")
            loComandoSeleccionar.AppendLine("		SET @lnEntero1 = @lnEntero1 + 1")
            loComandoSeleccionar.AppendLine("END;")
            loComandoSeleccionar.AppendLine("IF @lcAux <> '[ ]'         ")
            loComandoSeleccionar.AppendLine("   SET @lcCadena = @lcCadena + CHAR(13)+ N'SELECT Cod_Dep,         ")
            loComandoSeleccionar.AppendLine("												Nom_Dep,         ")
            loComandoSeleccionar.AppendLine("												Cantidad_Ventas,         ")
            loComandoSeleccionar.AppendLine("												Ventas_Totales,         ")
            loComandoSeleccionar.AppendLine("												Costo_Venta,         ")
            loComandoSeleccionar.AppendLine("												Tot_Net,         ")
            loComandoSeleccionar.AppendLine("												Por_Ven,         ")
            loComandoSeleccionar.AppendLine("												Total         ")
            loComandoSeleccionar.AppendLine("												'+@lcAux2+@lcAux4+',")
            loComandoSeleccionar.AppendLine("                                               Total*Por_Ven/100											    AS Total_Carga_Gasto,")
            loComandoSeleccionar.AppendLine("                                               Costo_Venta +  Total*Por_Ven/100                                AS Costo_Gasto,")
            loComandoSeleccionar.AppendLine("                                               Ventas_Totales - (Costo_Venta +  Total*Por_Ven/100)			    AS Utilidad,")
            loComandoSeleccionar.AppendLine("                                               (CASE Ventas_Totales - (Costo_Venta +  Total*Por_Ven/100)")
            loComandoSeleccionar.AppendLine("	                                                    WHEN 0 THEN 0")
            loComandoSeleccionar.AppendLine("	                                                    ELSE (Tot_Net/(Costo_Venta +  Total*Por_Ven/100))- 1")
            loComandoSeleccionar.AppendLine("                                               END)														    AS Porcentaje_Utilidad")
            loComandoSeleccionar.AppendLine("										FROM #TmpPrincipal          ")
            loComandoSeleccionar.AppendLine("											CROSS JOIN #tmpGastosTotales         ")
            loComandoSeleccionar.AppendLine("										ORDER BY Cod_Dep'         ")
            loComandoSeleccionar.AppendLine("Else         ")
            loComandoSeleccionar.AppendLine("	  SET @lcCadena = @lcCadena + CHAR(13)+ N'SELECT *, 0.00                                    AS Total,         ")
            loComandoSeleccionar.AppendLine("											            0.00                                    AS Total_Carga_Gasto,")
            loComandoSeleccionar.AppendLine("											            Ventas_Totales - Costo_Venta            AS Utilidad,")
            loComandoSeleccionar.AppendLine("											            Costo_Venta                             AS Costo_Gasto,")
            loComandoSeleccionar.AppendLine("                                                       (CASE Costo_Venta")
            loComandoSeleccionar.AppendLine("	                                                          WHEN 0 THEN 0")
            loComandoSeleccionar.AppendLine("	                                                          ELSE (Tot_Net/Costo_Venta)- 1")
            loComandoSeleccionar.AppendLine("                                                        END)									AS Porcentaje_Utilidad'+@lcAux4+'")
            loComandoSeleccionar.AppendLine("												FROM #TmpPrincipal             ")
            loComandoSeleccionar.AppendLine("												ORDER BY Cod_Dep'          ")

            loComandoSeleccionar.AppendLine("EXECUTE sp_executesql @lcCadena         ")


            loComandoSeleccionar.AppendLine("DROP TABLE #tmpGastos         ")
            loComandoSeleccionar.AppendLine("DROP TABLE #TmpPrincipal         ")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEstructura_Costos_Departamentos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrEstructura_Costos_Departamentos.ReportSource = loObjetoReporte

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
' EAG: 08/08/15: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' EAG: 01/09/15: Se cambio el codigo de cuenta de gasto por el nombre de cuenta de gasto	'
'-------------------------------------------------------------------------------------------'

