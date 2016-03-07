'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rOrdenes_Produccion_Russo"
'-------------------------------------------------------------------------------------------'
Partial Class rOrdenes_Produccion_Russo
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine("      DECLARE Cur CURSOR FOR    ")
            loComandoSeleccionar.AppendLine("          SELECT DISTINCT TOP 10 Talla FROM Ordenes_Produccion WHERE Talla <> ''       ")
            loComandoSeleccionar.AppendLine("       ")
            loComandoSeleccionar.AppendLine("      DECLARE		 @Temp NVARCHAR(MAX),        ")
            loComandoSeleccionar.AppendLine("   		             @Tallas NVARCHAR(MAX),        ")
            loComandoSeleccionar.AppendLine("   		             @TallasAux NVARCHAR(MAX),        ")
            loComandoSeleccionar.AppendLine("   		             @SubConsulta NVARCHAR(MAX),        ")
            loComandoSeleccionar.AppendLine("   		             @Numero INT        ")
            loComandoSeleccionar.AppendLine("      SET @SubConsulta = ''        ")
            loComandoSeleccionar.AppendLine("      SET @Tallas = ''        ")
            loComandoSeleccionar.AppendLine("      SET @TallasAux = ''        ")
            loComandoSeleccionar.AppendLine("      SET @Numero = 1        ")
            loComandoSeleccionar.AppendLine("      OPEN Cur        ")
            loComandoSeleccionar.AppendLine("      -- Consiguiendo todas las tallas        ")
            loComandoSeleccionar.AppendLine("          FETCH NEXT FROM Cur INTO @Temp        ")
            loComandoSeleccionar.AppendLine("          WHILE @@FETCH_STATUS = 0        ")
            loComandoSeleccionar.AppendLine("                  BEGIN        ")
            loComandoSeleccionar.AppendLine("                         SET @Tallas = @Tallas + '[' + RTrim(@Temp) + '],'        ")
            loComandoSeleccionar.AppendLine("                         SET @TallasAux = @TallasAux + 'ISNULL([' + RTrim(@Temp) + '],0) AS Columna'+CAST(@Numero AS NVARCHAR)+','        ")
            loComandoSeleccionar.AppendLine("                         SET @Numero = @Numero + 1        ")
            loComandoSeleccionar.AppendLine("                         FETCH NEXT FROM Cur INTO @Temp        ")
            loComandoSeleccionar.AppendLine("             End    ")
            loComandoSeleccionar.AppendLine("      CLOSE Cur        ")
            loComandoSeleccionar.AppendLine("      DEALLOCATE Cur        ")
            loComandoSeleccionar.AppendLine("      SET @Tallas = SUBSTRING(@Tallas, 0, LEN(@Tallas))        ")
            loComandoSeleccionar.AppendLine("      SET @TallasAux = SUBSTRING(@TallasAux, 0, LEN(@TallasAux))        ")
            loComandoSeleccionar.AppendLine("      -- Construyendo la consulta pivote        ")
            loComandoSeleccionar.AppendLine("      SET @SubConsulta = '    ")
            loComandoSeleccionar.AppendLine("      SELECT        ")
            loComandoSeleccionar.AppendLine("   	            Documento,    ")
            loComandoSeleccionar.AppendLine("   	            Talla,    ")
            loComandoSeleccionar.AppendLine("   	            Fec_Ini,    ")
            loComandoSeleccionar.AppendLine("   	            Modelo,    ")
            loComandoSeleccionar.AppendLine("   	            Cod_Art,    ")
            loComandoSeleccionar.AppendLine("   	            Nom_Art,    ")
            loComandoSeleccionar.AppendLine("   	            SUM(Can_Art1) AS Can_Art1    ")
            loComandoSeleccionar.AppendLine("      INTO #Temp         ")
            loComandoSeleccionar.AppendLine("      FROM		Ordenes_Produccion    ")
            loComandoSeleccionar.AppendLine("      WHERE   ")
            loComandoSeleccionar.AppendLine(" 			 Ordenes_Produccion.Documento between '" & lcParametro0Desde & "'")
            loComandoSeleccionar.AppendLine(" 			 AND '" & lcParametro0Hasta & "'")
            loComandoSeleccionar.AppendLine(" 			 AND Ordenes_Produccion.Fec_Ini between '" & lcParametro1Desde & "'")
            loComandoSeleccionar.AppendLine(" 			 AND '" & lcParametro1Hasta & "'")
            loComandoSeleccionar.AppendLine(" 			 AND Ordenes_Produccion.Cod_Art between '" & lcParametro2Desde & "'")
            loComandoSeleccionar.AppendLine(" 			 AND '" & lcParametro2Hasta & "'")
            loComandoSeleccionar.AppendLine("      GROUP BY Talla, Modelo, Cod_Art, Nom_Art, Fec_Ini, Documento    ")
            loComandoSeleccionar.AppendLine("      ORDER BY Talla, Modelo, Cod_Art, Nom_Art, Fec_Ini, Documento    ")
            loComandoSeleccionar.AppendLine("       ")
            loComandoSeleccionar.AppendLine("     SELECT Documento, Fec_Ini, Modelo, Cod_Art, Nom_Art,' + @TallasAux + '    ")
            loComandoSeleccionar.AppendLine("     FROM         ")
            loComandoSeleccionar.AppendLine("       (		         ")
            loComandoSeleccionar.AppendLine("             SELECT        ")
            loComandoSeleccionar.AppendLine("   	            #Temp.Documento,    ")
            loComandoSeleccionar.AppendLine("   	            RTRIM(#Temp.Talla) AS Talla,    ")
            loComandoSeleccionar.AppendLine("   	            #Temp.Fec_Ini,    ")
            loComandoSeleccionar.AppendLine("   	            #Temp.Modelo,    ")
            loComandoSeleccionar.AppendLine("   	            #Temp.Cod_Art,    ")
            loComandoSeleccionar.AppendLine("   	            #Temp.Nom_Art,    ")
            loComandoSeleccionar.AppendLine("   	            CAST(#Temp.Can_Art1 AS INT) AS Can_Art1    ")
            loComandoSeleccionar.AppendLine("             FROM  #Temp    ")
            loComandoSeleccionar.AppendLine("       ")
            loComandoSeleccionar.AppendLine("       ) AS SourceTable         ")
            loComandoSeleccionar.AppendLine("     PIVOT         ")
            loComandoSeleccionar.AppendLine("       (         ")
            loComandoSeleccionar.AppendLine("     SUM(Can_Art1)     ")
            loComandoSeleccionar.AppendLine("           FOR Talla IN ( '+ @Tallas +' )    ")
            loComandoSeleccionar.AppendLine("       ) AS PivotTable    ")
            loComandoSeleccionar.AppendLine("     ORDER BY " & lcOrdenamiento & ";'     ")
            loComandoSeleccionar.AppendLine("     EXEC sp_executesql  @SubConsulta        ")

            loComandoSeleccionar.AppendLine(" SELECT DISTINCT TOP 10 RTRIM(Talla) AS Talla FROM Ordenes_Produccion WHERE Talla <> '' ")

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rOrdenes_Produccion_Russo", laDatosReporte)



            Select Case laDatosReporte.Tables(1).Rows.Count
                Case 1
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
             
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "Replace (CStr({curReportes.Columna1}),"".00"" ,"""" )"

                Case 2
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "Replace (CStr({curReportes.Columna1}+{curReportes.Columna2}),"".00"" ,"""" )"

                Case 3
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "Replace (CStr({curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}),"".00"" ,"""" )"

                Case 4
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "Replace (CStr({curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}),"".00"" ,"""" )"
                Case 5
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text8"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text13"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "Replace (CStr({curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}),"".00"" ,"""" )"
                Case 6
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text8"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text10"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text13"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "Replace (CStr({curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}+{curReportes.Columna6}),"".00"" ,"""" )"
                Case 7
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text8"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text10"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text11"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text13"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(0).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "Replace (CStr({curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}+{curReportes.Columna6}+{curReportes.Columna7}),"".00"" ,"""" )"
                Case 8
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text8"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text10"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text11"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text12"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(7).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text13"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text2"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(7).Item(0).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "Replace (CStr({curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}+{curReportes.Columna6}+{curReportes.Columna7}+{curReportes.Columna8}),"".00"" ,"""" )"
                Case 9
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text8"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text10"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text11"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text12"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(7).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text14"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(8).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text13"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text2"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(7).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text21"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(8).Item(0).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "Replace (CStr({curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}+{curReportes.Columna6}+{curReportes.Columna7}+{curReportes.Columna8}+{curReportes.Columna9}),"".00"" ,"""" )"
                Case 10
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text8"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text10"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text11"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text12"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(7).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text14"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(8).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text15"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(9).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text13"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text2"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(7).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text21"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(8).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(9).Item(0).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}+{curReportes.Columna6}+{curReportes.Columna7}+{curReportes.Columna8}+{curReportes.Columna9}+{curReportes.Columna10}),"".00"" ,"""" )"
            End Select

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrOrdenes_Produccion_Russo.ReportSource = loObjetoReporte


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
' CMS:  07/07/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
