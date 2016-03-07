'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "gCobros_mRevision"
'-------------------------------------------------------------------------------------------'
Partial Class gCobros_mRevision
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()
            
            If cusAplicacion.goReportes.paParametrosIniciales(0) = 0 Then 
					lcParametro0Desde =  "'" & Now.Year & "'"
			End If
            
            loComandoSeleccionar.AppendLine("	SELECT 1 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total INTO #temp_cobros")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("	SELECT 2 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("	SELECT 3 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("	SELECT 4 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("	SELECT 5 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("	SELECT 6 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("	SELECT 7 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("	SELECT 8 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("	SELECT 9 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("	SELECT 10 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("	SELECT 11 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("	SELECT 12 AS Mes," & lcParametro0Desde & " AS Año, 'SIN REVISIÓN' AS Cod_Rev, 'SIN REVISIÓN' AS Nom_Rev, 'Debito    ' AS Tip_Doc, 0 AS Total")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("SELECT")
            loComandoSeleccionar.AppendLine("			DATEPART(MONTH, Cobros.fec_ini)						AS Mes,")
            loComandoSeleccionar.AppendLine("			DATEPART(YEAR, Cobros.fec_ini)						AS Año,")
            loComandoSeleccionar.AppendLine("  			ISNULL(Revisiones.Cod_Rev +'  ', 'SIN REVISIÓN')	AS Cod_Rev,")
            loComandoSeleccionar.AppendLine("  			ISNULL(RTRIM(Revisiones.Cod_Rev)+' - '+Revisiones.Nom_Rev, 'SIN REVISIÓN') AS Nom_Rev,")
            loComandoSeleccionar.AppendLine("  			Renglones_Cobros.Tip_Doc							AS Tip_Doc,")
            loComandoSeleccionar.AppendLine("  			SUM(Renglones_Cobros.Mon_Abo)						AS Total")
			loComandoSeleccionar.AppendLine("FROM		Cobros Cobros")
			'loComandoSeleccionar.AppendLine("	JOIN	Vendedores AS Vendedores ON  Vendedores.Cod_Ven = Cobros.Cod_Ven ")
			loComandoSeleccionar.AppendLine("	JOIN	Renglones_Cobros AS Renglones_Cobros ON Renglones_Cobros.Documento = Cobros.Documento")
			loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Cobrar AS Cuentas_Cobrar ON Cuentas_Cobrar.Documento = Renglones_Cobros.Doc_ori")
            loComandoSeleccionar.AppendLine(" 			AND	Cuentas_Cobrar.Cod_Tip = Renglones_Cobros.Cod_Tip")
			loComandoSeleccionar.AppendLine("	LEFT JOIN Revisiones AS Revisiones ON Revisiones.Cod_Rev = Cuentas_Cobrar.Cod_Rev")
            loComandoSeleccionar.AppendLine(" WHERE DATEPART(YEAR, Cobros.Fec_ini) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       AND Cobros.Cod_Suc BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("       AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY DATEPART(YEAR, Cobros.fec_ini), DATEPART(MONTH, Cobros.fec_ini), Revisiones.Cod_Rev, Revisiones.Nom_Rev, Renglones_Cobros.Tip_Doc")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  ")
            loComandoSeleccionar.AppendLine("   	#temp_cobros.Año		AS Año,  ")
            loComandoSeleccionar.AppendLine("   	#temp_cobros.Mes		AS Mes,  ")
            loComandoSeleccionar.AppendLine("   	#temp_cobros.Cod_Rev	AS Cod_Rev,  ")
            loComandoSeleccionar.AppendLine("   	#temp_cobros.Nom_Rev	AS Nom_Rev,  ")
            loComandoSeleccionar.AppendLine(" 		CASE")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 1 	THEN 'Ene'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 2 	THEN 'Feb'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 3 	THEN 'Mar'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 4 	THEN 'Abr'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 5 	THEN 'May'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 6 	THEN 'Jun'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 7 	THEN 'Jul'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 8 	THEN 'Ago'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 9 	THEN 'Sep'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 10	THEN 'Oct'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 11	THEN 'Nov'")
            loComandoSeleccionar.AppendLine(" 			WHEN #temp_cobros.Mes = 12	THEN 'Dic'")
            loComandoSeleccionar.AppendLine(" 		END						AS Str_Mes,")
            'loComandoSeleccionar.AppendLine("    	SUM(#temp_cobros.Total) AS Total")			   
            loComandoSeleccionar.AppendLine("    	SUM(")			   
            loComandoSeleccionar.AppendLine("    		CASE #temp_cobros.Tip_Doc ")			   
            loComandoSeleccionar.AppendLine("    			WHEN 'Debito' THEN Total")			   
            loComandoSeleccionar.AppendLine("    			WHEN 'Credito' THEN -Total")			   
            loComandoSeleccionar.AppendLine("    			ELSE  0")			   
            loComandoSeleccionar.AppendLine("    		END")			   
            loComandoSeleccionar.AppendLine("    	)AS Total")			   
            loComandoSeleccionar.AppendLine("    	")			   
            loComandoSeleccionar.AppendLine("    	")			   
            loComandoSeleccionar.AppendLine("    	")			   
            loComandoSeleccionar.AppendLine("    	")			   
            loComandoSeleccionar.AppendLine("FROM	#temp_cobros  ")
            loComandoSeleccionar.AppendLine("GROUP BY #temp_cobros.Año, #temp_cobros.Mes, #temp_cobros.Cod_Rev, #temp_cobros.Nom_Rev")
            loComandoSeleccionar.AppendLine("ORDER BY #temp_cobros.Año, #temp_cobros.Mes, #temp_cobros.Cod_Rev, #temp_cobros.Nom_Rev")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("gCobros_mRevision", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgCobros_mRevision.ReportSource = loObjetoReporte

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
' RJG: 23/08/10: Codigo inicial.															'
'-------------------------------------------------------------------------------------------'
