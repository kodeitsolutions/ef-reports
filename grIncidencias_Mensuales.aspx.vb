'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grIncidencias_Mensuales"
'-------------------------------------------------------------------------------------------'
Partial Class grIncidencias_Mensuales
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))

            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            
            If cusAplicacion.goReportes.paParametrosIniciales(0) = 0 Then
				
				lcParametro0Desde =  "'" & now.Year & "'"
				
			End If

            'If lcParametro3Desde = "" Then
            'lcParametro3Desde = ""
            'lcParametro3Hasta = "zzzzzzzz"
            'Else
            'lcParametro3Hasta = lcParametro3Desde
            'End If

            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            'If lcParametro4Desde = "" Then
            'lcParametro4Desde = ""
            'lcParametro4Hasta = "zzzzzzzz"
            'Else
            'lcParametro4Hasta = lcParametro4Desde
            'End If

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" DECLARE @Numero decimal")

            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(YEAR, Registro) As Año, ")
            loComandoSeleccionar.AppendLine(" 	        DATEPART(MONTH, Registro) AS Mes, ")
            loComandoSeleccionar.AppendLine(" 	        CASE  ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 1 THEN 'Ene' ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 2 THEN 'Feb' ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 3 THEN 'Mar' ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 4 THEN 'Abr' ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 5 THEN 'May' ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 6 THEN 'Jun' ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 7 THEN 'Jul' ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 8 THEN 'Ago' ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 9 THEN 'Sep' ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 10 THEN 'Oct' ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 11 THEN 'Nov' ")
            loComandoSeleccionar.AppendLine(" 		        WHEN DATEPART(MONTH, Registro) = 12 THEN 'Dic' ")
            loComandoSeleccionar.AppendLine(" 	        END AS Mes_Letras, ")
            loComandoSeleccionar.AppendLine("         	CASE  ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 1 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 2 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 3 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 4 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 5 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 6 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 7 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 8 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 9 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 10 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 11 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		WHEN DATEPART(MONTH, Registro) = 12 THEN SUM(Can_Err) ")
            loComandoSeleccionar.AppendLine("         		ELSE 0 ")
            loComandoSeleccionar.AppendLine("         	END AS Cantidad ")
            loComandoSeleccionar.AppendLine(" INTO      #Temp ")
            loComandoSeleccionar.AppendLine(" FROM      Factory_Global.dbo.Errores ")
            loComandoSeleccionar.AppendLine(" WHERE     DATEPART(YEAR, Registro) = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND Cod_Usu             Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Status              IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine(" 			AND Sistema             Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Modulo              Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY  DATEPART(YEAR, Registro), DATEPART(MONTH, Registro) ")
            loComandoSeleccionar.AppendLine(" ORDER BY  1,2 ")
            'loComandoSeleccionar.AppendLine(" ORDER BY  DATEPART(YEAR, Registro), DATEPART(MONTH, Registro) ")

            'Me.Response.Clear()
            'Me.Response.ContentType = "text/plain"
            'Me.Response.Write(loComandoSeleccionar.ToString())
            'Me.Response.Flush()
            'Me.Response.End()
            'Return
            ' Desde aqui se comienza a agrupar mes por mes el reporte
            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	1 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'Ene' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")
            loComandoSeleccionar.AppendLine(" INTO #Temp2")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	2 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'Feb' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	3 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'Mar' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	4 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'Abr' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	5 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'May' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	6 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'Jun' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	7 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'Jul' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	8 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'Ago' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	9 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'Sep' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	10 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'Oct' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	11 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'Nov' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	" & lcParametro0Desde & " AS Año,")
            loComandoSeleccionar.AppendLine(" 	12 AS Mes,")
            loComandoSeleccionar.AppendLine(" 	'Dic' AS Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	0 AS Cantidad")

            loComandoSeleccionar.AppendLine(" UNION ALL")

            loComandoSeleccionar.AppendLine(" SELECT ")
            loComandoSeleccionar.AppendLine(" 	Año,")
            loComandoSeleccionar.AppendLine(" 	Mes,")
            loComandoSeleccionar.AppendLine(" 	Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	Cantidad")
            loComandoSeleccionar.AppendLine(" FROM #Temp")


            loComandoSeleccionar.AppendLine(" SELECT    Año,")
            loComandoSeleccionar.AppendLine(" 	        Mes,")
            loComandoSeleccionar.AppendLine(" 	        Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	        SUM(Cantidad) AS Cantidad,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 1 THEN SUM(Cantidad) ELSE 0 END AS Enero,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 2 THEN SUM(Cantidad) ELSE 0 END AS Febrero,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 3 THEN SUM(Cantidad) ELSE 0 END AS Marzo,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 4 THEN SUM(Cantidad) ELSE 0 END AS Abril,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 5 THEN SUM(Cantidad) ELSE 0 END AS Mayo,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 6 THEN SUM(Cantidad) ELSE 0 END AS Junio,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 7 THEN SUM(Cantidad) ELSE 0 END AS Julio,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 8 THEN SUM(Cantidad) ELSE 0 END AS Agosto,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 9 THEN SUM(Cantidad) ELSE 0 END AS Septiembre,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 10 THEN SUM(Cantidad) ELSE 0 END AS Octubre,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 11 THEN SUM(Cantidad) ELSE 0 END AS Noviembre,")
            loComandoSeleccionar.AppendLine("           CASE  WHEN Mes = 12 THEN SUM(Cantidad) ELSE 0 END AS Diciembre")
            loComandoSeleccionar.AppendLine(" FROM      #Temp2")
            loComandoSeleccionar.AppendLine(" GROUP BY  Año,Mes,Mes_Letras")
            loComandoSeleccionar.AppendLine(" ORDER BY  Año,Mes,Mes_Letras")

            loComandoSeleccionar.AppendLine(" SET @Numero = (SELECT MAX(Cantidad) FROM #Temp2)")

            loComandoSeleccionar.AppendLine(" SELECT    CAST((ROUND(@Numero, -(LEN(CAST(@Numero As varchar))-4))*0.2*1) AS DECIMAL) AS E1,")
            loComandoSeleccionar.AppendLine(" 		    CAST((ROUND(@Numero, -(LEN(CAST(@Numero As varchar))-4))*0.2*2) AS DECIMAL) AS E2,")
            loComandoSeleccionar.AppendLine(" 		    CAST((ROUND(@Numero, -(LEN(CAST(@Numero As varchar))-4))*0.2*3) AS DECIMAL) AS E3,")
            loComandoSeleccionar.AppendLine(" 		    CAST((ROUND(@Numero, -(LEN(CAST(@Numero As varchar))-4))*0.2*4) AS DECIMAL) AS E4,")
            loComandoSeleccionar.AppendLine(" 		    CAST((ROUND(@Numero, -(LEN(CAST(@Numero As varchar))-4))*0.2*5) AS DECIMAL) AS E5,")
            loComandoSeleccionar.AppendLine(" 		    CAST((ROUND(@Numero, -(LEN(CAST(@Numero As varchar))-4))*0.2*6) AS DECIMAL) AS E6")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grIncidencias_Mensuales", laDatosReporte)

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Or (laDatosReporte.Tables(1).Rows(0).Item(5).ToString = "0") Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte.DataDefinition.FormulaFields("E1").Text = "ToText(Replace (ToText(" & laDatosReporte.Tables(1).Rows(0).Item(0) & "), '.00' , '' ))"
            loObjetoReporte.DataDefinition.FormulaFields("E1").Text = "Replace (ToText(" & loObjetoReporte.DataDefinition.FormulaFields("E1").Text & "), ',' , '.' )"

            loObjetoReporte.DataDefinition.FormulaFields("E2").Text = "ToText(Replace (ToText(" & laDatosReporte.Tables(1).Rows(0).Item(1) & "), '.00' , '' ))"
            loObjetoReporte.DataDefinition.FormulaFields("E2").Text = "Replace (ToText(" & loObjetoReporte.DataDefinition.FormulaFields("E2").Text & "), ',' , '.' )"

            loObjetoReporte.DataDefinition.FormulaFields("E3").Text = "ToText(Replace (ToText(" & laDatosReporte.Tables(1).Rows(0).Item(2) & "), '.00' , '' ))"
            loObjetoReporte.DataDefinition.FormulaFields("E3").Text = "Replace (ToText(" & loObjetoReporte.DataDefinition.FormulaFields("E3").Text & "), ',' , '.' )"

            loObjetoReporte.DataDefinition.FormulaFields("E4").Text = "ToText(Replace (ToText(" & laDatosReporte.Tables(1).Rows(0).Item(3) & "), '.00' , '' ))"
            loObjetoReporte.DataDefinition.FormulaFields("E4").Text = "Replace (ToText(" & loObjetoReporte.DataDefinition.FormulaFields("E4").Text & "), ',' , '.' )"

            loObjetoReporte.DataDefinition.FormulaFields("E5").Text = "ToText(Replace (ToText(" & laDatosReporte.Tables(1).Rows(0).Item(4) & "), '.00' , '' ))"
            loObjetoReporte.DataDefinition.FormulaFields("E5").Text = "Replace (ToText(" & loObjetoReporte.DataDefinition.FormulaFields("E5").Text & "), ',' , '.' )"

            loObjetoReporte.DataDefinition.FormulaFields("E6").Text = "ToText(Replace (ToText(" & laDatosReporte.Tables(1).Rows(0).Item(5) & "), '.00' , '' ))"
            loObjetoReporte.DataDefinition.FormulaFields("E6").Text = "Replace (ToText(" & loObjetoReporte.DataDefinition.FormulaFields("E6").Text & "), ',' , '.' )"


            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvgrIncidencias_Mensuales.ReportSource = loObjetoReporte

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
' JJD: 07/11/09: Codigo inicial
'-------------------------------------------------------------------------------------------'