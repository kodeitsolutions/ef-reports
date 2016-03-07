'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArticulos_sAlmacenConsolidado_OnyxTD"
'-------------------------------------------------------------------------------------------'
Partial Class rArticulos_sAlmacenConsolidado_OnyxTD
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
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
            Dim lcParametro10Desde As String = cusAplicacion.goReportes.paParametrosIniciales(10)
            Dim lcParametro11Desde As String = cusAplicacion.goReportes.paParametrosIniciales(11)
            Dim lcExisiencia As String = ""

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()
            Dim lcExistencia As String = ""

            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpExistencias (  Cod_Art   VARCHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Nom_Art   VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Cod_Uni1  VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Exi_Act1  DECIMAL(28,10),")
            loConsulta.AppendLine("                                Cod_Alm   VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Nom_Alm   VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Cod_Dep   VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Cod_Sec   VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Cod_Tip   VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Cod_Cla   VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Cod_Mar   VARCHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Upc       VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Modelo    VARCHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Cos_Pro1  DECIMAL(28,10),")
            loConsulta.AppendLine("                                Cantidad_Caja  DECIMAL(28,10))")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpExistencias(Cod_Art, Nom_Art, Cod_Uni1, Exi_Act1, ")
            loConsulta.AppendLine("                            Cod_Alm, Nom_Alm, Cod_Dep, Cod_Sec, Cod_Tip,")
            loConsulta.AppendLine("                            Cod_Cla, Cod_Mar, Upc, Modelo, Cos_Pro1, Cantidad_Caja)")
            loConsulta.AppendLine("SELECT	")
            loConsulta.AppendLine("        Articulos.Cod_Art, ")
            loConsulta.AppendLine("        Articulos.Nom_Art, ")
            loConsulta.AppendLine("        Articulos.Cod_Uni1,")

			Select Case lcParametro11Desde
                Case "Actual"
                    loConsulta.AppendLine("        COALESCE(Renglones_Almacenes.Exi_Act1,0)     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Act1"
                Case "Comprometida"
                    loConsulta.AppendLine("        COALESCE(Renglones_Almacenes.Exi_Ped1,0)     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Ped1"
                Case "Cotizada"
                    loConsulta.AppendLine("        COALESCE(Renglones_Almacenes.Exi_Cot1,0)     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Cot1"
                Case "En_Produccion"
                    loConsulta.AppendLine("        COALESCE(Renglones_Almacenes.Exi_Pro1,0)     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Pro1"
                Case "Por_Llegar"
                    loConsulta.AppendLine("        COALESCE(Renglones_Almacenes.Exi_Por1,0)     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Por1"
                Case "Por_Despachar"
                    loConsulta.AppendLine("        COALESCE(Renglones_Almacenes.Exi_Des1,0)     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Des1"
                Case "Por_Distribuir"
                    loConsulta.AppendLine("        COALESCE(Renglones_Almacenes.Exi_Dis1,0)     AS Exi_Act1,")
                    lcExistencia = "Renglones_Almacenes.Exi_Dis1"
                Case Else
                    loConsulta.AppendLine("        0     AS Exi_Act1")
                    lcExistencia = "Renglones_Almacenes.Exi_Act1"
            End Select					 			
			
            loConsulta.AppendLine("        Almacenes.Cod_Alm, ")
            loConsulta.AppendLine("        Almacenes.Nom_Alm,")
            loConsulta.AppendLine("        Articulos.Cod_Dep,")
            loConsulta.AppendLine("        Articulos.Cod_Sec,")
            loConsulta.AppendLine("        Articulos.Cod_Tip,")
            loConsulta.AppendLine("        Articulos.Cod_Cla,")
            loConsulta.AppendLine("        Articulos.Cod_Mar,")
            loConsulta.AppendLine("        Articulos.Upc,")
            loConsulta.AppendLine("        Articulos.Modelo,")
            loConsulta.AppendLine("        Articulos.Cos_Pro1,")
            loConsulta.AppendLine("        CAST(1 AS DECIMAL(28, 10)) AS Cantidad_Caja")
            loConsulta.AppendLine("FROM    Articulos")
            loConsulta.AppendLine("    JOIN Renglones_Almacenes")
            loConsulta.AppendLine("        ON Renglones_Almacenes.Cod_Art = Articulos.Cod_Art")
            loConsulta.AppendLine("    JOIN Almacenes")
            loConsulta.AppendLine("        ON Almacenes.Cod_Alm = Renglones_Almacenes.Cod_Alm")
            loConsulta.AppendLine(" WHERE		Articulos.Cod_Art BETWEEN " & lcParametro0Desde & "	AND " & lcParametro0Hasta)
            loConsulta.AppendLine(" 		AND Articulos.status IN (" & lcParametro1Desde & ")")
            loConsulta.AppendLine(" 		AND Renglones_Almacenes.Cod_Alm BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loConsulta.AppendLine(" 		AND Articulos.Cod_Dep BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine(" 		AND Articulos.Cod_Sec BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loConsulta.AppendLine(" 		AND Articulos.Cod_Mar BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loConsulta.AppendLine(" 		AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loConsulta.AppendLine(" 		AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loConsulta.AppendLine("      	AND Articulos.Cod_Ubi BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loConsulta.AppendLine("      	AND Articulos.Cod_Pro BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta)

            
            Select Case lcParametro10Desde
                Case "Todos"
                    loConsulta.AppendLine("      ")
                Case "Igual"
                    loConsulta.AppendLine("        AND " & lcExistencia & "          =   0  ")
                Case "Mayor"							  
                    loConsulta.AppendLine("        AND " & lcExistencia & "          >   0  ")
                Case "Menor"							  
                    loConsulta.AppendLine("        AND " & lcExistencia & "          <   0  ")
                Case "Maximo"							  
                    loConsulta.AppendLine("        AND Articulos.Exi_Max           =   " & lcExistencia & "  ")
                Case "Minimo"							  
                    loConsulta.AppendLine("        AND Articulos.Exi_Min           =   " & lcExistencia & "  ")
                Case "Pedido"
                    loConsulta.AppendLine("        AND Articulos.Exi_pto           =   " & lcExistencia & "  ")
            End Select
            
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Transpone la existencia por almacén para convertir en columnas")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DECLARE @lcColumnas AS VARCHAR(MAX);")
            loConsulta.AppendLine("DECLARE @lcSumas AS VARCHAR(MAX);")
            loConsulta.AppendLine("DECLARE @lcConsulta AS VARCHAR(MAX);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT @lcColumnas = STUFF((SELECT distinct ',' + QUOTENAME(RTRIM(Cod_Alm)) ")
            loConsulta.AppendLine("                    from #tmpExistencias")
            loConsulta.AppendLine("                    FOR XML PATH(''), TYPE")
            loConsulta.AppendLine("                    ).value('.', 'NVARCHAR(MAX)'), 1, 1, '');")
            loConsulta.AppendLine("                    ")
            loConsulta.AppendLine("SELECT @lcSumas = STUFF((SELECT distinct ',COALESCE(SUM(' + QUOTENAME(RTRIM(Cod_Alm))+'),0) AS ' + QUOTENAME(RTRIM(Cod_Alm))")
            loConsulta.AppendLine("                    from #tmpExistencias")
            loConsulta.AppendLine("                    FOR XML PATH(''), TYPE")
            loConsulta.AppendLine("                    ).value('.', 'NVARCHAR(MAX)'), 1, 1, '');")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SET @lcConsulta =  'SELECT   Cod_Art, Nom_Art, Cod_Uni1, Cod_Dep, Cod_Sec, Cod_Tip,")
            loConsulta.AppendLine("                            Cod_Cla, Cod_Mar, Upc, Modelo, Cos_Pro1, Cantidad_Caja, ' + @lcSumas + ' ")
            loConsulta.AppendLine("                    FROM    (SELECT * FROM #tmpExistencias) AS Datos")
            loConsulta.AppendLine("                    PIVOT   (SUM(Exi_Act1) FOR Cod_Alm IN (' + @lcColumnas + ')) AS P ")
            loConsulta.AppendLine("                    GROUP BY Cod_Art, Nom_Art, Cod_Uni1, Cod_Dep, Cod_Sec, Cod_Tip, ")
            loConsulta.AppendLine("                             Cod_Cla, Cod_Mar, Upc, Modelo, Cos_Pro1, Cantidad_Caja")
            loConsulta.AppendLine("                    ';")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("EXEC(@lcConsulta);")
            loConsulta.AppendLine("--DROP TABLE #tmpExistencias;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            

            Dim loServicios As New cusDatos.goDatos
            
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArticulos_sAlmacenConsolidado_OnyxTD", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArticulos_sAlmacenConsolidado_OnyxTD.ReportSource = loObjetoReporte

            'Selección de opcion por excel (Microsoft Excel - xls)
            If (Me.Request.QueryString("salida").ToLower() = "xls") Then
                ' Ruta donde se creara temporalmente el archivo
                Dim lcFileName As String = Server.MapPath("~\Administrativo\Temporales\rArticulos_sAlmacenConsolidado_OnyxTD_" & Guid.NewGuid().ToString("N") & ".xls")
                ' Se exporta para crear el archivo temporal
                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, lcFileName)

                ' Se modifica el contenido del archivo
                Dim lcParametrosReporte As String = cusAplicacion.goReportes.mObtenerParametros(cusAplicacion.goReportes.paNombresParametros, cusAplicacion.goReportes.paParametrosIniciales, cusAplicacion.goReportes.paParametrosFinales)
                Me.mGenerarArchivoExcel(lcFileName, laDatosReporte.Tables(0), lcParametrosReporte)

                ' Se coloca en la respuesta para decargar
                Me.Response.Clear()
                Me.Response.Buffer = True 
                Me.Response.AppendHeader("content-disposition", "attachment; filename=rArticulos_sAlmacenConsolidado_OnyxTD.xls")
                Me.Response.ContentType = "application/excel"
                Me.Response.WriteFile(lcFileName, True)
                Me.Response.Write(Space(30))
                Me.Response.Flush()
                Me.Response.Close()
                
				Me.Response.End()

            Else

                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "Este reporte fue diseñado solo para mostrar la vista ""Microsoft Excel - xls"". ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")

            End If

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

	''' <summary>
	''' Genera la salida personalizada en Excel. 
	''' </summary>
	''' <param name="lcNombreArchivo">Nombre del archivo a generar.</param>
	''' <param name="loDatos">Tabla de datos a generar.</param>
	''' <param name="lcParametrosReporte">Cadena con los parámetros originales del reporte.</param>
	''' <remarks></remarks>
	Private Sub mGenerarArchivoExcel(ByVal lcNombreArchivo As String, ByVal loDatos As DataTable, ByVal lcParametrosReporte As String)
		
		Dim lnDecimalesMonto As Integer = goOpciones.pnDecimalesParaMonto
		Dim lnDecimalesCosto As Integer = goOpciones.pnDecimalesParaCosto
		Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
		Dim lnDecimalesPorcentaje As Integer = goOpciones.pnDecimalesParaPorcentaje
		
		Dim lcFormatoMontos As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesMonto)
		Dim lcFormatoCosto As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesCosto)
		
		Dim lcFormatoCantidad As String 
		If (lnDecimalesCantidad > 0) Then 
			lcFormatoCantidad = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesCantidad)
		Else
			lcFormatoCantidad = "###,###,###,###,##0"
		End If
		
		Dim lcFormatoPorcentaje As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesPorcentaje)
        
        Dim laAlmacenes As New Generic.Dictionary(Of String, String)
        Dim laColumasAlm() As String = New String(){"H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", _
            "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", _
            "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", _
            "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", _
            "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ"}
        

	 '******************************************************************'
	 ' Declaración de objetos de excel: IMPORTANTE liberar recursos al	'
	 ' final usando el GARBAGE COLLECTOR y ReleaseComObject.			'
	 '******************************************************************'
		Dim loExcel		As Excel.Application	= Nothing
		Dim laLibros	As Excel.Workbooks		= Nothing
		Dim loLibro		As Excel.Workbook		= Nothing
    	Dim loHoja		As Excel.Worksheet		= Nothing
		Dim loCeldas	As Excel.Range			= Nothing
		Dim loRango		As Excel.Range			= Nothing
		
		Dim loFilas		As Excel.Range			= Nothing
		Dim loColumnas	As Excel.Range			= Nothing
		Dim loFormas	As Excel.Shapes			= Nothing
		Dim loImagen	As Excel.Shape			= Nothing
		Dim loFuente	As Excel.Font			= Nothing
 
		
		
        Try
        
        ' Se inicializa el objeto de aplicacion excel
            loExcel = New Excel.Application()
            loExcel.Visible = False
            loExcel.DisplayAlerts = False 

        ' Crea un nuevo libro de excel y activa la primera hoja
            laLibros = loExcel.Workbooks
            'loLibro = laLibros.Add()
            
            'Dim lcPlantilla As String = HttpContext.Current.Server.MapPath("~/Administrativo/Complementos/plantilla.xls")
            'System.IO.File.Copy(lcPlantilla, lcNombreArchivo)
            loLibro = laLibros.Open(lcNombreArchivo)
            
            loHoja = loLibro.Worksheets(1)
            loHoja.Activate()

		' Formato por defecto de todas las celdas			
			loCeldas = loHoja.Range("A1:IV65536")
            'loCeldas = loHoja.Cells
			loCeldas.Clear()
            loFuente = loCeldas.Font
            loFuente.Size = 9
            loFuente.Name = "Tahoma"


		 '******************************************************************'
		 ' Encabezado de la hoja											'
		 '******************************************************************'
			'Dim lcLogo As String = goEmpresa.pcUrlLogo 
			'lcLogo = HttpContext.Current.Server.MapPath(lcLogo)
			'loFormas = loHoja.Shapes

			'loFormas.AddPicture(lcLogo,  Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 1, 1, 60, 60)
			
            loRango = loHoja.Range("A1")
            loRango.Value = cusAplicacion.goEmpresa.pcNombre
            loFuente = loRango.Font
            loFuente.Size = 12
            loFuente.Bold = True
            
            loRango = loHoja.Range("A2")
            loRango.Value = cusAplicacion.goEmpresa.pcRifEmpresa
            loFuente = loRango.Font
            loFuente.Size = 12
            loFuente.Bold = True

            loRango = loHoja.Range("A3")
            loRango.Value = goReportes.pcModuloReporte
            loFuente = loRango.Font
            loFuente.Bold = True

            loRango = loHoja.Range("B4:L4")
            loRango.Select()
            loRango.MergeCells = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.Value = "Reporte de Artículos con su Stock por Almacén Consolidado (ONYXTD)"
            loFuente = loRango.Font
            loFuente.Size = 14
            loFuente.Bold = True

            ' Parametros del reporte
            loRango = loHoja.Range("B5:L5")
            loRango.Select()
            loRango.MergeCells = True
            loRango.Value = lcParametrosReporte
            loRango.WrapText = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loFuente = loRango.Font
            loFuente.Bold = True

            ' Fecha y hora de creacion
			Dim ldFecha As DateTime = Date.Now()
			loRango = loHoja.Range("L1")
			loRango.NumberFormat = "mm/dd/yyyy;@"
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
			loRango.Value = ldFecha
			
			loRango = loHoja.Range("L2")
			loRango.NumberFormat = "@" 'La celda almacena un string
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
			loRango.Value = ldFecha.ToString("hh:mm:ss tt")


			Dim lnFilaActual As Integer = 7

		 '******************************************************************'
		 ' Datos del Reporte												'
		 '******************************************************************'

			loRango = loHoja.Range("B" & lnFilaActual)
			loRango.Value = "Código"
			
			loRango = loHoja.Range("C" & lnFilaActual)
			loRango.Value = "Artículo"
			
			loRango = loHoja.Range("D" & lnFilaActual)
			loRango.Value = "Unidad"
			
			loRango = loHoja.Range("E" & lnFilaActual)
			loRango.Value = "UPC"
			
			loRango = loHoja.Range("F" & lnFilaActual)
			loRango.Value = "Modelo"
			
			loRango = loHoja.Range("G" & lnFilaActual)
			loRango.Value = "Costo Promedio"
			
			'loRango = loHoja.Range("H" & lnFilaActual)
			'loRango.Value = "Unid. x Cajas"
			
            ' Posición (base 0) de la columna del primer almacén en la tabla
            Const KC_PRIMER_ALMACEN As Integer= 12
            Dim lcUltimaColumna As String = "G"
            ' ALMACENES 
			For n AS Integer = KC_PRIMER_ALMACEN To loDatos.Columns.Count - 1 	

                laAlmacenes.Add(laColumasAlm(n - KC_PRIMER_ALMACEN), loDatos.Columns(n).ColumnName)
                lcUltimaColumna = laColumasAlm(n - KC_PRIMER_ALMACEN)
			
			    loRango = loHoja.Range(lcUltimaColumna & lnFilaActual)
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
				loRango.NumberFormat = "@"
			    loRango.Value = loDatos.Columns(n).ColumnName

            Next n
			
			loRango = loHoja.Range("B" & lnFilaActual & ":" & lcUltimaColumna & lnFilaActual)
			loFuente = loRango.Font
			loFuente.Bold = True
			loFuente.Color = Rgb(255, 255, 255)
			loRango.Interior.Color = Rgb(0, 51, 153)
			
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
			loRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, RGB(255,255,255))
			loRango.Borders( Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous 
			loRango.Borders( Excel.XlBordersIndex.xlInsideVertical).Weight = Excel.XlBorderWeight.xlMedium 
			loRango.Borders( Excel.XlBordersIndex.xlInsideVertical).Color = RGB(255,255,255)
			
			Dim lnFilaInicio As Integer  = lnFilaActual
			For lnRenglon As Integer = 0 To loDatos.Rows.Count - 1
				Dim loRenglon As DataRow = loDatos.Rows(lnRenglon)
				
				lnFilaActual += 1
			
				'Código
				loRango = loHoja.Range("B" & lnFilaActual)
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Cod_Art")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'Artículo
				loRango = loHoja.Range("C" & lnFilaActual)
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Nom_Art")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'Unidad
				loRango = loHoja.Range("D" & lnFilaActual)
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Cod_Uni1")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'UPC 
				loRango = loHoja.Range("E" & lnFilaActual)
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Upc")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'Modelo
				loRango = loHoja.Range("F" & lnFilaActual)	
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Modelo")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
						
				'Costo Promedio
				loRango = loHoja.Range("G" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoCosto
				loRango.Value = CDec(loRenglon("Cos_Pro1"))
					
				''Unid. x Cajas
				'loRango = loHoja.Range("H" & lnFilaActual)
				'loRango.NumberFormat = "###,###,###,###,##0.000"
				'loRango.Value = CDec(loRenglon("Cantidad_Caja")) 
				
                For Each lcColumna As String In laAlmacenes.Keys
                    
				    'Existencia en el almacén
				    loRango = loHoja.Range(lcColumna & lnFilaActual) 
				    loRango.NumberFormat = lcFormatoCantidad
				    loRango.Value = goServicios.mRedondearValor(CDec(loRenglon(laAlmacenes(lcColumna))), lnDecimalesMonto)

                Next lcColumna
				
			Next lnRenglon
			
			Dim lnTotal As Integer = loDatos.Rows.Count
			loRango = loHoja.Range("B" & (lnFilaInicio) & ":" & lcUltimaColumna & (lnFilaInicio))
			loRango.Select() 
			'loExcel.Selection.AutoFilter()
			
			loRango = loHoja.Range("B" & (lnFilaInicio+1) & ":" & lcUltimaColumna & (lnFilaInicio + lnTotal))
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
			
			Dim lnDesde AS Integer = lnFilaInicio
			Dim lnHasta AS Integer = lnFilaInicio + lnTotal
			
			'lnFilaInicio += lnTotal + 2
			loRango = loHoja.Range("B" & (lnHasta + 1))
			loRango.NumberFormat = "@"
			loRango.Value = "Total Artículos: " & lnTotal.ToString()
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            
			loRango = loHoja.Range("B" & (lnFilaInicio) & ":" & lcUltimaColumna & (lnFilaInicio))
			loFuente = loRango.Font
			loFuente.Bold = True
					
			loFilas = loCeldas.Rows
			loFilas.AutoFit()
			
			loRango = loHoja.Range("B1:B" & lnFilaInicio)
			loRango.ColumnWidth = 20
			
			loRango = loHoja.Range("C1:C" & lnFilaInicio)
			loRango.ColumnWidth = 40
			
			loRango = loHoja.Range("D1:D" & lnFilaInicio)
			loRango.ColumnWidth = 8
			
			loRango = loHoja.Range("E1:E" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("F1:F" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("G1:G" & lnFilaInicio)
			loRango.ColumnWidth = 14
			
			loRango = loHoja.Range("H1:H" & lnFilaInicio)
			loRango.ColumnWidth = 13
			
			loRango = loHoja.Range("I1:" & lcUltimaColumna & lnFilaInicio)
			loRango.ColumnWidth = 12
						
            ' Seleccionamos la primera celda del libro
			loRango = loHoja.Range("A1")
            loRango.Select()

            'Guardamos los cambios del libro activo
            loLibro.SaveAs(lcNombreArchivo)
            
		 '******************************************************************'
		 ' IMPORTANTE: Forma correcta de liberar recursos!!!				'
		 '******************************************************************'
            ' Cerramos y liberamos recursos

        Catch loExcepcion As Exception
			
			Throw New Exception("No fue posible exportar los datos a excel. " & loExcepcion.Message, loExcepcion)
			
        Finally

			If (loFuente IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loFuente)
				loFuente = Nothing
			End If
			
			If (loFormas IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loFormas)
				loFormas = Nothing
			End If
			
			If (loRango IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loRango)
				loRango = Nothing
			End If
			
			If (loFilas IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loFilas)
				loFilas = Nothing
			End If
			
			If (loColumnas IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loColumnas)
				loColumnas = Nothing
			End If
			
			If (loCeldas IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loCeldas)
				loCeldas = Nothing
			End If
			
			If (loHoja IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loHoja)
				loHoja = Nothing
			End If
			
			If (loLibro IsNot Nothing) Then
				loLibro.Close(True)
				System.Runtime.InteropServices.Marshal.ReleaseComObject(loLibro)
				loLibro = Nothing
			End If

			If (laLibros IsNot Nothing) Then
				System.Runtime.InteropServices.Marshal.ReleaseComObject(laLibros)
				laLibros = Nothing
			End If
            
            loExcel.Quit()

			System.Runtime.InteropServices.Marshal.ReleaseComObject(loExcel)
            loExcel = Nothing 
            
            GC.Collect()
            GC.WaitForPendingFinalizers()
            
        End Try

    End Sub


End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 26/09/13: Codigo inicial, a partir de rArticulos_sAlmacenConsolidado.                '
'-------------------------------------------------------------------------------------------'
