'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTVentas_ArticulosSucursal"
'-------------------------------------------------------------------------------------------'
Partial Class rTVentas_ArticulosSucursal

    Inherits vis2formularios.frmReporte

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
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim leParametro12Desde As Integer = cusAplicacion.goReportes.paParametrosIniciales(12)
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
            Dim lcParametro14Desde As Integer = cusAplicacion.goReportes.paParametrosIniciales(14)
            Dim lcParametro15Desde As String = cusAplicacion.goReportes.paParametrosIniciales(15)

            Dim lcSql As String

            If leParametro12Desde > 0 Then
                lcSql = "Select Top " + leParametro12Desde.ToString
            Else
                lcSql = "Select "
            End If

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            '-------------------------------------------------------------------------------------------'
            ' Seleccion de las sucursales
            '-------------------------------------------------------------------------------------------'
            If lcParametro14Desde = 0 Then
                lcParametro14Desde = 7
            End If

            If Not (lcParametro15Desde = "") Then
                'lcParametro15Desde = LTrim(RTrim(lcParametro15Desde & "'"))
                lcParametro15Desde = lcParametro15Desde.Replace(",,,,", "")
                lcParametro15Desde = lcParametro15Desde.Replace(",,,", "")
                lcParametro15Desde = lcParametro15Desde.Replace(",,", "")
                lcParametro15Desde = lcParametro15Desde.Replace(" ,  ", "")
                'lcParametro15Desde = LTrim(RTrim(lcParametro15Desde & "'"))

                If lcParametro15Desde.Substring(lcParametro15Desde.Length - 1) = "," Then
                    lcParametro15Desde = lcParametro15Desde.Remove(lcParametro15Desde.Length - 1)
                End If

                If lcParametro15Desde.Substring(0) = "," Then
                    lcParametro15Desde = lcParametro15Desde.Remove(0)
                End If
            End If

            lcParametro15Desde = lcParametro15Desde.Replace("`", "'")
            lcParametro15Desde = lcParametro15Desde.Replace("´", "'")
            lcParametro15Desde = lcParametro15Desde.Replace(",", "','")
            lcParametro15Desde = lcParametro15Desde.Replace(" ", "")

            loComandoSeleccionar.AppendLine(" DECLARE curSucursales CURSOR FOR    ")
            If lcParametro15Desde = "" Then
                loComandoSeleccionar.AppendLine(" SELECT DISTINCT TOP " & lcParametro14Desde & " Nom_Suc FROM Sucursales   ")
            Else
                loComandoSeleccionar.AppendLine(" SELECT DISTINCT TOP " & lcParametro14Desde & " Nom_Suc FROM Sucursales WHERE Cod_Suc IN ('" & lcParametro15Desde & "')    ")
            End If

            loComandoSeleccionar.AppendLine("     ")
            loComandoSeleccionar.AppendLine(" DECLARE	@tmpSucursales NVARCHAR(MAX),    ")
            loComandoSeleccionar.AppendLine("			@Sucursales NVARCHAR(MAX),    ")
            loComandoSeleccionar.AppendLine("			@SucursalesAux NVARCHAR(MAX),    ")
            loComandoSeleccionar.AppendLine("			@SucursalesAux2 NVARCHAR(MAX),    ")
            loComandoSeleccionar.AppendLine("			@SubConsulta NVARCHAR(MAX),    ")
            loComandoSeleccionar.AppendLine("			@Numero  INT,    ")
            loComandoSeleccionar.AppendLine("			@Numero2 INT    ")

            loComandoSeleccionar.AppendLine(" SET @SubConsulta		= ''   ")
            loComandoSeleccionar.AppendLine(" SET @Sucursales		= ''   ")
            loComandoSeleccionar.AppendLine(" SET @SucursalesAux	= ''   ")
            loComandoSeleccionar.AppendLine(" SET @SucursalesAux2	= ''   ")
            loComandoSeleccionar.AppendLine(" SET @Numero			= 1    ")
            loComandoSeleccionar.AppendLine(" SET @Numero2			= 8    ")
            loComandoSeleccionar.AppendLine(" OPEN curSucursales  ")
            loComandoSeleccionar.AppendLine(" -- Consiguiendo todas las sucursales    ")
            loComandoSeleccionar.AppendLine(" FETCH NEXT FROM curSucursales INTO @tmpSucursales    ")
            loComandoSeleccionar.AppendLine("	WHILE @@FETCH_STATUS = 0    ")
            loComandoSeleccionar.AppendLine("		BEGIN ")
            loComandoSeleccionar.AppendLine("			SET @Sucursales		=	@Sucursales + '[' + RTrim(@tmpSucursales) + '],'    ")
            loComandoSeleccionar.AppendLine("			SET @SucursalesAux	=	@SucursalesAux + 'ISNULL([' + RTrim(@tmpSucursales) + '],0) AS Columna' + CAST(@Numero AS NVARCHAR) + ',' ")
            loComandoSeleccionar.AppendLine("			SET @SucursalesAux2	=	@SucursalesAux2 + 'ISNULL([' + RTrim(@tmpSucursales) + '],0) AS Columna' + CAST(@Numero2 AS NVARCHAR) + ',' ")
            loComandoSeleccionar.AppendLine("			SET @Numero			=	@Numero + 1    ")
            loComandoSeleccionar.AppendLine("			SET @Numero2		=	@Numero2 + 1    ")
            loComandoSeleccionar.AppendLine("			FETCH NEXT FROM curSucursales INTO @tmpSucursales    ")
            loComandoSeleccionar.AppendLine("		END ")
            loComandoSeleccionar.AppendLine(" CLOSE curSucursales ")
            loComandoSeleccionar.AppendLine(" DEALLOCATE curSucursales ")
            loComandoSeleccionar.AppendLine(" SET @Sucursales		= SUBSTRING(@Sucursales, 0, LEN(@Sucursales))    ")
            loComandoSeleccionar.AppendLine(" SET @SucursalesAux	= SUBSTRING(@SucursalesAux, 0, LEN(@SucursalesAux))    ")
            loComandoSeleccionar.AppendLine(" SET @SucursalesAux2	= SUBSTRING(@SucursalesAux2, 0, LEN(@SucursalesAux2))    ")

            loComandoSeleccionar.AppendLine(" SET @SubConsulta	= '	")
            loComandoSeleccionar.AppendLine(" SELECT    Articulos.Cod_Art				AS	Cod_Art, ")
            loComandoSeleccionar.AppendLine("           Articulos.Nom_Art			    AS	Nom_Art, ")
            loComandoSeleccionar.AppendLine("           SUM(Renglones_Facturas.Can_Art1)	    AS  Can_Art, ")
            loComandoSeleccionar.AppendLine("           SUM(Renglones_Facturas.Mon_Net)		AS  Mon_Net, ")
            loComandoSeleccionar.AppendLine("           LTRIM(Sucursales.Nom_Suc)		AS	Nom_Suc ")
            loComandoSeleccionar.AppendLine(" INTO		#curTemporal ")
            loComandoSeleccionar.AppendLine(" FROM		Facturas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Facturas, ")
            loComandoSeleccionar.AppendLine("           Articulos, ")
            'loComandoSeleccionar.AppendLine("           Vendedores, ")
            'loComandoSeleccionar.AppendLine("           Clientes, ")
            'loComandoSeleccionar.AppendLine("           Departamentos, ")
            'loComandoSeleccionar.AppendLine("           Monedas, ")
            'loComandoSeleccionar.AppendLine("           Secciones, ")
            'loComandoSeleccionar.AppendLine("           Clases_articulos, ")
            'loComandoSeleccionar.AppendLine("           Tipos_Articulos, ")
            'loComandoSeleccionar.AppendLine("           Proveedores, ")
            loComandoSeleccionar.AppendLine("           Sucursales ")
            loComandoSeleccionar.AppendLine(" WHERE		Facturas.Documento				 =   Renglones_Facturas.Documento ")
            loComandoSeleccionar.AppendLine("			AND Sucursales.Cod_Suc			 =	Facturas.Cod_Suc ")
            'loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Cli             =   Clientes.Cod_Cli ")
            'loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Ven             =   Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           AND Renglones_Facturas.Cod_Art   =   Articulos.Cod_Art ")
            'loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep            =   Departamentos.Cod_Dep ")
            'loComandoSeleccionar.AppendLine("           AND Monedas.Cod_Mon				 =	Clientes.Cod_Mon ")
            'loComandoSeleccionar.AppendLine("           AND Secciones.Cod_Dep			 =	Departamentos.Cod_Dep ")
            'loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Sec			 =	Secciones.Cod_Sec ")
            'loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Cla			 =	Clases_Articulos.Cod_Cla ")
            'loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Tip			 =	Tipos_Articulos.Cod_Tip ")
            'loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Pro			 =	Proveedores.Cod_Pro ")
            'loComandoSeleccionar.AppendLine("           AND Renglones_Facturas.Cod_Art	 =	Articulos.Cod_Art ")
            loComandoSeleccionar.AppendLine("           AND Facturas.Fec_Ini			 BETWEEN '" & lcParametro0Desde & "'")
            loComandoSeleccionar.AppendLine("           AND '" & lcParametro0Hasta & "'")
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Cli			 BETWEEN '" & lcParametro1Desde & "'")
            loComandoSeleccionar.AppendLine("           AND '" & lcParametro1Hasta & "'")
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Ven			 BETWEEN '" & lcParametro2Desde & "'")
            loComandoSeleccionar.AppendLine("           AND '" & lcParametro2Hasta & "'")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Dep			 BETWEEN '" & lcParametro3Desde & "'")
            loComandoSeleccionar.AppendLine("           AND '" & lcParametro3Hasta & "'")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Sec			 BETWEEN '" & lcParametro4Desde & "'")
            loComandoSeleccionar.AppendLine("           AND '" & lcParametro4Hasta & "'")
            loComandoSeleccionar.AppendLine("           AND Renglones_Facturas.Cod_Art	 BETWEEN '" & lcParametro5Desde & "'")
            loComandoSeleccionar.AppendLine("           AND '" & lcParametro5Hasta & "'")
            loComandoSeleccionar.AppendLine("           AND Renglones_Facturas.Cod_Alm	 BETWEEN '" & lcParametro6Desde & "'")
            loComandoSeleccionar.AppendLine("           AND '" & lcParametro6Hasta & "'")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Cla			 BETWEEN '" & lcParametro7Desde & "'")
            loComandoSeleccionar.AppendLine("           AND '" & lcParametro7Hasta & "'")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Tip			 BETWEEN '" & lcParametro8Desde & "'")
            loComandoSeleccionar.AppendLine("           AND '" & lcParametro8Hasta & "'")
            loComandoSeleccionar.AppendLine("           AND Articulos.Cod_Pro			 BETWEEN '" & lcParametro9Desde & "'")
            loComandoSeleccionar.AppendLine("           AND '" & lcParametro9Hasta & "'")
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Mon				 BETWEEN '" & lcParametro10Desde & "'")
            loComandoSeleccionar.AppendLine("           AND '" & lcParametro10Hasta & "'")
            loComandoSeleccionar.AppendLine("           AND Facturas.Status IN (" & lcParametro11Desde.Replace("'", "''") & ")")
            loComandoSeleccionar.AppendLine("           AND Facturas.Cod_Rev			 BETWEEN '" & lcParametro13Desde & "'")
            loComandoSeleccionar.AppendLine("    	    AND '" & lcParametro13Hasta & "'")
            loComandoSeleccionar.AppendLine(" GROUP BY Articulos.Cod_Art, Articulos.Nom_Art, Facturas.Cod_Suc, LTRIM(Sucursales.Nom_Suc)")

            '-------------------------------------------------------------------------------------------'
            ' Ejecucion del PIVOT											 IN (' + @Sucursales + '))	    
            '-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine(" SELECT Cod_Art + '': '' + Nom_Art AS Articulo, Cod_Art, ' + @SucursalesAux + ' ")
            loComandoSeleccionar.AppendLine(" INTO #Temp")
            loComandoSeleccionar.AppendLine(" FROM   ")
            loComandoSeleccionar.AppendLine(" (SELECT RTRIM(Cod_Art) AS Cod_Art, Nom_Art, Nom_Suc, Can_Art, Mon_Net FROM #curTemporal) AS SourceTable")
            loComandoSeleccionar.AppendLine(" PIVOT (SUM(Can_Art) FOR Nom_Suc IN (' + @Sucursales + ')) AS PivotTable ")

            loComandoSeleccionar.AppendLine(" SELECT Cod_Art + '': '' + Nom_Art AS Articulo, Cod_Art, ' + @SucursalesAux2 + ' ")
            loComandoSeleccionar.AppendLine(" INTO #Temp2")
            loComandoSeleccionar.AppendLine(" FROM   ")
            loComandoSeleccionar.AppendLine(" (SELECT RTRIM(Cod_Art) AS Cod_Art, Nom_Art, Nom_Suc, Can_Art, Mon_Net FROM #curTemporal) AS SourceTable")
            loComandoSeleccionar.AppendLine(" PIVOT (SUM(Mon_Net) FOR Nom_Suc IN (' + @Sucursales + ')) AS PivotTable ")

            loComandoSeleccionar.AppendLine(" SELECT * FROM #Temp")
            loComandoSeleccionar.AppendLine(" JOIN #Temp2 ON #Temp.Cod_Art = #Temp2.Cod_Art")
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)

            loComandoSeleccionar.AppendLine(" '")
            loComandoSeleccionar.AppendLine(" EXEC sp_ExecuteSQL  @SubConsulta ")


            If lcParametro15Desde = "" Then
                loComandoSeleccionar.AppendLine(" SELECT DISTINCT TOP " & lcParametro14Desde & " RTRIM(Nom_Suc) AS Nom_Suc, RTRIM(Cod_Suc) AS Cod_Suc FROM Sucursales    ")
            Else
                loComandoSeleccionar.AppendLine(" SELECT DISTINCT TOP " & lcParametro14Desde & " RTRIM(Nom_Suc) AS Nom_Suc, RTRIM(Cod_Suc) AS Cod_Suc FROM Sucursales WHERE Cod_Suc IN ('" & lcParametro15Desde & "')    ")
            End If

           

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTVentas_ArticulosSucursal", laDatosReporte)

            Dim Tope As Integer
            Tope = laDatosReporte.Tables(1).Rows.Count

            If Tope < lcParametro14Desde Then
                lcParametro14Desde = Tope
            End If

            Select Case lcParametro14Desde
                Case 1
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}"
                    loObjetoReporte.DataDefinition.FormulaFields("Total_Monto").Text = "{curReportes.Columna8}"
                Case 2
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}"
                    loObjetoReporte.DataDefinition.FormulaFields("Total_Monto").Text = "{curReportes.Columna8}+{curReportes.Columna9}"
                Case 3
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}"
                    loObjetoReporte.DataDefinition.FormulaFields("Total_Monto").Text = "{curReportes.Columna8}+{curReportes.Columna9}+{curReportes.Columna10}"
                Case 4
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}"
                    loObjetoReporte.DataDefinition.FormulaFields("Total_Monto").Text = "{curReportes.Columna8}+{curReportes.Columna9}+{curReportes.Columna10}+{curReportes.Columna11}"
                Case 5
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}"
                    loObjetoReporte.DataDefinition.FormulaFields("Total_Monto").Text = "{curReportes.Columna8}+{curReportes.Columna9}+{curReportes.Columna10}+{curReportes.Columna11}+{curReportes.Columna12}"
                Case 6
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text15"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}+{curReportes.Columna6}"
                    loObjetoReporte.DataDefinition.FormulaFields("Total_Monto").Text = "{curReportes.Columna8}+{curReportes.Columna9}+{curReportes.Columna10}+{curReportes.Columna11}+{curReportes.Columna12}+{curReportes.Columna13}"
                Case 7
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text9"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text3"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text4"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text5"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text6"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text7"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(0).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text8"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(0).ToString

                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text20"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(0).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text19"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(1).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text18"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(2).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text17"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(3).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text16"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(4).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text15"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(5).Item(1).ToString
                    CType(loObjetoReporte.ReportDefinition.ReportObjects("text14"), CrystalDecisions.CrystalReports.Engine.TextObject).Text = laDatosReporte.Tables(1).Rows(6).Item(1).ToString

                    loObjetoReporte.DataDefinition.FormulaFields("Total").Text = "{curReportes.Columna1}+{curReportes.Columna2}+{curReportes.Columna3}+{curReportes.Columna4}+{curReportes.Columna5}+{curReportes.Columna6}+{curReportes.Columna7}"
                    loObjetoReporte.DataDefinition.FormulaFields("Total_Monto").Text = "{curReportes.Columna8}+{curReportes.Columna9}+{curReportes.Columna10}+{curReportes.Columna11}+{curReportes.Columna12}+{curReportes.Columna13}+{curReportes.Columna14}"

            End Select

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTVentas_ArticulosSucursal.ReportSource = loObjetoReporte

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
' CMS: 30/07/09: Programacion inicial
'-------------------------------------------------------------------------------------------'
