'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "grPedidos_MensualesCantidades"
'-------------------------------------------------------------------------------------------'
Partial Class grPedidos_MensualesCantidades
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
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
            Dim lcParametro7Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = cusAplicacion.goReportes.paParametrosIniciales(9)
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine(" DECLARE @Numero decimal")

            loComandoSeleccionar.AppendLine(" SELECT    DATEPART(YEAR, Pedidos.Fec_Ini) As Año, ")
            loComandoSeleccionar.AppendLine(" 	        DATEPART(MONTH, Pedidos.Fec_Ini) AS Mes, ")

            loComandoSeleccionar.AppendLine(" 	CASE  ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 1 THEN 'Ene' ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 2 THEN 'Feb' ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 3 THEN 'Mar' ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 4 THEN 'Abr' ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 5 THEN 'May' ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 6 THEN 'Jun' ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 7 THEN 'Jul' ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 8 THEN 'Ago' ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 9 THEN 'Sep' ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 10 THEN 'Oct' ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 11 THEN 'Nov' ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 12 THEN 'Dic' ")
            loComandoSeleccionar.AppendLine(" 	END AS Mes_Letras, ")
            loComandoSeleccionar.AppendLine(" 	CASE  ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 1 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 2 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 3 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 4 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 5 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 6 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 7 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 8 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 9 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 10 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 11 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		WHEN DATEPART(MONTH, Pedidos.Fec_Ini) = 12 THEN SUM(Renglones_Pedidos.Can_Art1) ")
            loComandoSeleccionar.AppendLine(" 		ELSE 0 ")
            loComandoSeleccionar.AppendLine(" 	END AS Cantidad ")
            loComandoSeleccionar.AppendLine(" INTO #Temp ")
            loComandoSeleccionar.AppendLine(" FROM Pedidos ")
            loComandoSeleccionar.AppendLine(" JOIN Renglones_Pedidos ON Pedidos.Documento   =   Renglones_Pedidos.Documento ")
            loComandoSeleccionar.AppendLine(" JOIN Articulos ON Articulos.Cod_Art           =   Renglones_Pedidos.Cod_Art ")

            loComandoSeleccionar.AppendLine(" WHERE         ")
            loComandoSeleccionar.AppendLine("      			DATEPART(YEAR, Pedidos.Fec_Ini) =   " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 				AND Renglones_Pedidos.Cod_Art   Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Pedidos.Cod_Cli             Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Pedidos.Cod_Ven             Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Dep           Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Tip           Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Articulos.Cod_Cla           Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Pedidos.Status              IN (" & lcParametro7Desde & ")")

            If lcParametro9Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 			AND Pedidos.Cod_Rev between " & lcParametro8Desde)
            Else
                loComandoSeleccionar.AppendLine(" 			AND Pedidos.Cod_Rev NOT between " & lcParametro8Desde)
            End If
            loComandoSeleccionar.AppendLine(" 				AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("               AND Pedidos.Cod_Suc between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("               AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine(" GROUP BY DATEPART(YEAR, Pedidos.Fec_Ini), DATEPART(MONTH, Pedidos.Fec_Ini) ")
            loComandoSeleccionar.AppendLine(" ORDER BY DATEPART(YEAR, Pedidos.Fec_Ini), DATEPART(MONTH, Pedidos.Fec_Ini) ")




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

            loComandoSeleccionar.AppendLine(" SELECT    Año,")
            loComandoSeleccionar.AppendLine(" 	        Mes,")
            loComandoSeleccionar.AppendLine(" 	        Mes_Letras,")
            loComandoSeleccionar.AppendLine(" 	        Cantidad")
            loComandoSeleccionar.AppendLine(" FROM      #Temp")


            loComandoSeleccionar.AppendLine(" SELECT    Año,")
            loComandoSeleccionar.AppendLine(" 	        Mes,")
            loComandoSeleccionar.AppendLine(" 	        Mes_Letras,")
            loComandoSeleccionar.AppendLine("           SUM(Cantidad) AS Cantidad,")

            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 1  THEN SUM(Cantidad) ELSE 0 END AS Enero,")
            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 2  THEN SUM(Cantidad) ELSE 0 END AS Febrero,")
            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 3  THEN SUM(Cantidad) ELSE 0 END AS Marzo,")
            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 4  THEN SUM(Cantidad) ELSE 0 END AS Abril,")
            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 5  THEN SUM(Cantidad) ELSE 0 END AS Mayo,")
            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 6  THEN SUM(Cantidad) ELSE 0 END AS Junio,")
            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 7  THEN SUM(Cantidad) ELSE 0 END AS Julio,")
            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 8  THEN SUM(Cantidad) ELSE 0 END AS Agosto,")
            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 9  THEN SUM(Cantidad) ELSE 0 END AS Septiembre,")
            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 10 THEN SUM(Cantidad) ELSE 0 END AS Octubre,")
            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 11 THEN SUM(Cantidad) ELSE 0 END AS Noviembre,")
            loComandoSeleccionar.AppendLine(" CASE WHEN Mes = 12 THEN SUM(Cantidad) ELSE 0 END AS Diciembre")


            loComandoSeleccionar.AppendLine(" FROM #Temp2")

            loComandoSeleccionar.AppendLine(" GROUP BY Año,Mes,Mes_Letras")
            loComandoSeleccionar.AppendLine(" ORDER BY Año,Mes,Mes_Letras")

            loComandoSeleccionar.AppendLine(" SET @Numero = (SELECT MAX(Cantidad) FROM #Temp2)")

            loComandoSeleccionar.AppendLine(" Select	Cast((round(@Numero, -(len(Cast(@Numero As varchar))-4))*0.2*1) AS DECIMAL) AS E1,")
            loComandoSeleccionar.AppendLine(" 		    Cast((round(@Numero, -(len(Cast(@Numero As varchar))-4))*0.2*2) AS DECIMAL) AS E2,")
            loComandoSeleccionar.AppendLine(" 		    Cast((round(@Numero, -(len(Cast(@Numero As varchar))-4))*0.2*3) AS DECIMAL) AS E3,")
            loComandoSeleccionar.AppendLine(" 		    Cast((round(@Numero, -(len(Cast(@Numero As varchar))-4))*0.2*4) AS DECIMAL) AS E4,")
            loComandoSeleccionar.AppendLine(" 		    Cast((round(@Numero, -(len(Cast(@Numero As varchar))-4))*0.2*5) AS DECIMAL) AS E5,")
            loComandoSeleccionar.AppendLine(" 		    Cast((round(@Numero, -(len(Cast(@Numero As varchar))-4))*0.2*6) AS DECIMAL) AS E6")



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("grPedidos_MensualesCantidades", laDatosReporte)

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
            Me.crvgrPedidos_MensualesCantidades.ReportSource = loObjetoReporte

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
' JJD: 22/08/09: Codigo inicial
'-------------------------------------------------------------------------------------------'