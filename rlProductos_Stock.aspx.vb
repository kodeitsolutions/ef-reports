'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rlProductos_Stock"
'-------------------------------------------------------------------------------------------'
Partial Class rlProductos_Stock
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
		

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            
            'Precio/Costo
            Dim lcParametro7Desde As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(7)).ToUpper().Trim()
            'Moneda
            Dim lcParametro8Desde As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(8)).Trim()
            'Tasa
            Dim lnParametro9Desde As String = CDec(cusAplicacion.goReportes.paParametrosIniciales(9))
            'Actual/Disponible
            Dim lcParametro10Desde As String  = CStr(cusAplicacion.goReportes.paParametrosIniciales(10)).ToUpper().Trim()
            'Nivel de Stock
            Dim lcParametro11Desde As String = CStr(cusAplicacion.goReportes.paParametrosIniciales(11)).ToUpper().Trim()
            
		'-------------------------------------------------------------------------------------------'
        ' Precio/Costo a mostrar																	
		'-------------------------------------------------------------------------------------------'
            Dim lcMonto As String
            Dim lcCampoMoneda_Base As String
            Dim lcTituloMonto As String
            Select Case lcParametro7Desde
				Case "PRECIO1"
					lcMonto = "Articulos.Precio1"
					lcTituloMonto = "Precio 1"
					lcCampoMoneda_Base = "Articulos.Pre_Nac"
				Case "PRECIO2"
					lcMonto = "Articulos.Precio2"
					lcTituloMonto = "Precio 2"
					lcCampoMoneda_Base = "Articulos.Pre_Nac"
				Case "PRECIO3"
					lcMonto = "Articulos.Precio3"
					lcTituloMonto = "Precio 3"
					lcCampoMoneda_Base = "Articulos.Pre_Nac"
				Case "PRECIO4"
					lcMonto = "Articulos.Precio4"
					lcTituloMonto = "Precio 4"
					lcCampoMoneda_Base = "Articulos.Pre_Nac"
				Case "PRECIO5"
					lcMonto = "Articulos.Precio5"
					lcTituloMonto = "Precio 5"
					lcCampoMoneda_Base = "Articulos.Pre_Nac"
				Case "ULTIMO_MN"	
					lcMonto = "Articulos.Cos_Ult1"
					lcTituloMonto = "Costo Último"
					lcCampoMoneda_Base = "CAST(1 AS BIT)"
				Case "ULTIMO_OM"	
					lcMonto = "Articulos.Cos_Ult2"
					lcTituloMonto = "Costo Último OM"
					lcCampoMoneda_Base = "CAST(0 AS BIT)"
				Case "PROMEDIO_MN"	
					lcMonto = "Articulos.Cos_Pro1"
					lcTituloMonto = "Costo Promedio"
					lcCampoMoneda_Base = "CAST(1 AS BIT)"
				Case "PROMEDIO_OM"	
					lcMonto = "Articulos.Cos_Pro2"
					lcTituloMonto = "Costo Promedio OM"
					lcCampoMoneda_Base = "CAST(0 AS BIT)"
				Case "ANTERIOR_MN"	
					lcMonto = "Articulos.Cos_Ant1"
					lcTituloMonto = "Costo Anterior"
					lcCampoMoneda_Base = "CAST(1 AS BIT)"
				Case "ANTERIOR_OM"	
					lcMonto = "Articulos.Cos_Ant2"
					lcTituloMonto = "Costo Anterior OM"
					lcCampoMoneda_Base = "CAST(0 AS BIT)"
				Case "CLIENTE_MN"	
					lcMonto = "Articulos.Cod_Cli1"
					lcTituloMonto = "Costo Cliente"
					lcCampoMoneda_Base = "CAST(1 AS BIT)"
				Case "CLIENTE_OM"	
					lcMonto = "Articulos.Cos_Cli2"
					lcTituloMonto = "Costo Cliente OM"
					lcCampoMoneda_Base = "CAST(0 AS BIT)"
				Case Else
					lcMonto = "Articulos.Precio1"
					lcTituloMonto = "Precio 1"
					lcCampoMoneda_Base = "Articulos.Pre_Nac"
			End Select
			lcTituloMonto = goServicios.mObtenerCampoFormatoSQL(lcTituloMonto)

			
		'-------------------------------------------------------------------------------------------'
        ' Moneda y tasa a mostrar																	
		'-------------------------------------------------------------------------------------------'
            Dim llMostrarMonedaBase As Boolean = False
            Dim llMostrarMonedaAdicional As Boolean = False
            
			If (lcParametro8Desde = "") OrElse (lcParametro8Desde=goMoneda.pcCodigoMonedaBase) Then 
				llMostrarMonedaBase = True
			ElseIf (lcParametro8Desde=goMoneda.pcCodigoMonedaAdicional) Then 
				llMostrarMonedaAdicional = True
			End If
			
			Dim lnTasaAdicional As Decimal = goMoneda.pnTasaMonedaAdicional
			Dim lnTasaOtra As Decimal = lnParametro9Desde
			
			If (lnTasaOtra <= 0D) Then
				
				If llMostrarMonedaBase Then 
					lnTasaOtra = 1D
				ElseIf llMostrarMonedaAdicional Then 
					lnTasaOtra = goMoneda.pnTasaMonedaAdicional 
				Else
					lnTasaOtra = goMoneda.mObtenerValorTasa(lcParametro8Desde, Date.Now())
				End If
			
			End If
			
			'Si se desea mostrar en tasa adicional: usar la tasa adicional 
			'indicada en lugar de la actual
			If llMostrarMonedaAdicional Then 
				lnTasaAdicional = lnTasaOtra 
			End If
			
		'-------------------------------------------------------------------------------------------'
        ' Tipo de inventario a mostrar																	
		'-------------------------------------------------------------------------------------------'
			Dim lcStock As String 
            Dim lcTituloStock As String
			If lcParametro10Desde = "ACTUAL" Then
				lcStock = "Articulos.Exi_Act1"
				lcTituloStock = "'Actual'"
			Else
				lcStock = "(Articulos.Exi_Act1 - Articulos.Exi_Ped1)"
				lcTituloStock = "'Disponible'"
			End If
			
		'-------------------------------------------------------------------------------------------'
        ' Tipo de inventairo a mostrar																	
		'-------------------------------------------------------------------------------------------'
			Dim lcCondicionStock As String 
			
			Select Case lcParametro11Desde
				Case "MAXIMO"
					lcCondicionStock = lcStock & " = Articulos.Exi_Max" 
				Case "MINIMO"
					lcCondicionStock = lcStock & " = Articulos.Exi_Min" 
				Case "PEDIDO"
					lcCondicionStock = lcStock & " = Articulos.Exi_Pto" 
				Case "MAYOR"
					lcCondicionStock = lcStock & " > 0" 
				Case "MENOR"
					lcCondicionStock = lcStock & " < 0" 
				Case "IGUAL"
					lcCondicionStock = lcStock & " = 0" 
				Case Else
					lcCondicionStock = ""
			End Select
			

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("DECLARE @lnTasaAdicional DECIMAL(28, 10);")
            loConsulta.AppendLine("DECLARE @lnTasaOtra DECIMAL(28, 10);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SET @lnTasaAdicional = " & goServicios.mObtenerCampoFormatoSQL(lnTasaAdicional) & ";")
            loConsulta.AppendLine("SET @lnTasaOtra = " & goServicios.mObtenerCampoFormatoSQL(lnTasaOtra) & ";")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpProductos( Cod_Art CHAR(30) COLLATE Database_Default,")
            loConsulta.AppendLine("							Nom_Art CHAR(100) COLLATE Database_Default,")
            loConsulta.AppendLine("							Modelo CHAR(30) COLLATE Database_Default,")
            loConsulta.AppendLine("							Moneda_Base BIT,")
            loConsulta.AppendLine("							Exi_Act1 DECIMAL(28, 10),")
            loConsulta.AppendLine("							Exi_Dis1 DECIMAL(28, 10),")
            loConsulta.AppendLine("							Monto DECIMAL(28, 10),")
            loConsulta.AppendLine("							Stock DECIMAL(28, 10),")
            loConsulta.AppendLine("							Titulo_Monto CHAR(50) COLLATE Database_Default,")
            loConsulta.AppendLine("							Titulo_Stock CHAR(50) COLLATE Database_Default,")
            loConsulta.AppendLine("							)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO	#tmpProductos")
            loConsulta.AppendLine("		(	Cod_Art, Nom_Art, Modelo, Moneda_Base, ")
            loConsulta.AppendLine("			Exi_Act1, Exi_Dis1, Monto, Stock, ")
            loConsulta.AppendLine("			Titulo_Monto, Titulo_Stock)")
            loConsulta.AppendLine("SELECT	Articulos.Cod_Art		AS Cod_Art,")
            loConsulta.AppendLine("			Articulos.Nom_Art		AS Nom_Art,")
            loConsulta.AppendLine("			Articulos.Modelo		AS Modelo,")
            loConsulta.AppendLine("			" & lcCampoMoneda_Base & "		AS Moneda_Base,")
            loConsulta.AppendLine("			Articulos.Exi_Act1		AS Exi_Act1,")
            loConsulta.AppendLine("			(Articulos.Exi_Act1-Articulos.Exi_Ped1)		AS Exi_Dis1,")
            loConsulta.AppendLine("			" & lcMonto & "			AS Monto,")
            loConsulta.AppendLine("			" & lcStock & "			AS Stock,")
            loConsulta.AppendLine("			" & lcTituloMonto & "			AS Titulo_Monto,")
            loConsulta.AppendLine("			" & lcTituloStock & "			AS Titulo_Stock")
            loConsulta.AppendLine("FROM		Articulos ")
            loConsulta.AppendLine("WHERE		Articulos.Cod_Art		BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("			AND " & lcParametro0Hasta)
            loConsulta.AppendLine("		AND Articulos.Nom_Art		BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("			AND " & lcParametro1Hasta)
            loConsulta.AppendLine("		AND Articulos.Status        IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine("		AND Articulos.Cod_Dep		BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("			AND " & lcParametro3Hasta)
            loConsulta.AppendLine("		AND Articulos.Cod_Sec		BETWEEN " & lcParametro4Desde)
            loConsulta.AppendLine("			AND " & lcParametro4Hasta)
            loConsulta.AppendLine("		AND Articulos.Cod_Mar		BETWEEN " & lcParametro5Desde)
            loConsulta.AppendLine("			AND " & lcParametro5Hasta)
            loConsulta.AppendLine("		AND Articulos.Cod_Pro		BETWEEN " & lcParametro6Desde)
            loConsulta.AppendLine("			AND " & lcParametro6Hasta)
            If (lcCondicionStock<>"") Then	
				loConsulta.AppendLine("			AND " & lcCondicionStock)
			End If
            loConsulta.AppendLine("UPDATE	#tmpProductos")
            loConsulta.AppendLine("SET		Monto = CASE")
            If llMostrarMonedaBase Then
				loConsulta.AppendLine("			WHEN (Moneda_Base = 1) ")
				loConsulta.AppendLine("				THEN Monto")
				loConsulta.AppendLine("			WHEN (Moneda_Base = 0) ")
				loConsulta.AppendLine("				THEN ROUND(Monto*@lnTasaAdicional, " & goOpciones.pnDecimalesParaMonto & ")")
			ElseIf llMostrarMonedaAdicional Then
				loConsulta.AppendLine("			WHEN (Moneda_Base = 1) ")
				loConsulta.AppendLine("				THEN ROUND(Monto/@lnTasaAdicional, " & goOpciones.pnDecimalesParaMonto & ")")
				loConsulta.AppendLine("			WHEN (Moneda_Base = 0) ")
				loConsulta.AppendLine("				THEN Monto")
            Else
				loConsulta.AppendLine("			WHEN (Moneda_Base = 1) ")
				loConsulta.AppendLine("				THEN ROUND(Monto/@lnTasaOtra, " & goOpciones.pnDecimalesParaMonto & ")")
				loConsulta.AppendLine("			WHEN (Moneda_Base = 0) ")
				loConsulta.AppendLine("				THEN ROUND(Monto*@lnTasaAdicional/@lnTasaOtra, " & goOpciones.pnDecimalesParaMonto & ")")
            End If
            loConsulta.AppendLine("				END")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Cod_Art, Nom_Art, Modelo, Moneda_Base, ")
            loConsulta.AppendLine("			Exi_Act1, Exi_Dis1, Monto, Stock, ")
            loConsulta.AppendLine("			Titulo_Monto, Titulo_Stock")
            loConsulta.AppendLine("FROM		#tmpProductos")
            loConsulta.AppendLine("ORDER BY      " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DROP TABLE #tmpProductos")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos
			'Me.mEscribirConsulta(loConsulta.ToString)
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

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


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rlProductos_Stock", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrlProductos_Stock.ReportSource = loObjetoReporte

            'Selección de opcion por excel (Microsoft Excel - xls)
            If (Me.Request.QueryString("salida").ToLower() = "xls") Then
                ' Ruta donde se creara temporalmente el archivo
                Dim lcFileName As String = Server.MapPath("~\Administrativo\Temporales\rlProductos_Stock_" & Guid.NewGuid().ToString("N") & ".xls")
                ' Se exporta para crear el archivo temporal
                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, lcFileName)

                ' Se modifica el contenido del archivo
                Dim lcParametrosReporte As String = cusAplicacion.goReportes.mObtenerParametros(cusAplicacion.goReportes.paNombresParametros, cusAplicacion.goReportes.paParametrosIniciales, cusAplicacion.goReportes.paParametrosFinales)
                Me.mGenerarArchivoExcel(lcFileName, laDatosReporte.Tables(0), lcParametrosReporte, lcTituloMonto.Replace("'", ""))

                ' Se coloca en la respuesta para decargar
                Me.Response.Clear()
                Me.Response.Buffer = True 
                Me.Response.AppendHeader("content-disposition", "attachment; filename=rlProductos_Stock.xls")
                Me.Response.ContentType = "application/excel"
                Me.Response.WriteFile(lcFileName, True)
                Me.Response.Write(Space(30))
                Me.Response.Flush()
                Me.Response.Close()
                
				Me.Response.End()
                
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
''' <param name="lcTituloCampoMonto">(parámetro adicional) Título del campo de monto mostrado</param>
''' <remarks></remarks>
	Private Sub mGenerarArchivoExcel(ByVal lcNombreArchivo As String, ByVal loDatos As DataTable, ByVal lcParametrosReporte As String, ByVal lcTituloCampoMonto As String)
		
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
            
            loRango = loHoja.Range("A2")
            loRango.Value = cusAplicacion.goEmpresa.pcRifEmpresa

            loRango = loHoja.Range("B5:G5")
            loRango.Select()
            loRango.MergeCells = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.Value = "Listado de Productos en Stock al " & (Date.Today()).ToString("MM/dd/yyyy")
            loFuente = loRango.Font
            loFuente.Size = 14
            loFuente.Bold = True

            ' Fecha y hora de creacion
			Dim ldFecha As DateTime = Date.Now()
			loRango = loHoja.Range("G1")
			loRango.NumberFormat = "mm/dd/yyyy;@"
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
			loRango.Value = ldFecha
			
			loRango = loHoja.Range("G2")
			loRango.NumberFormat = "@" 'La celda almacena un string
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
			loRango.Value = ldFecha.ToString("hh:mm:ss tt")

            ' Parametros del reporte
            loRango = loHoja.Range("B7:G7")
            loRango.Select()
            loRango.MergeCells = True
            loRango.Value = lcParametrosReporte
            loRango.WrapText = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify


			Dim lnFilaActual As Integer = 9

		 '******************************************************************'
		 ' Datos del Reporte												'
		 '******************************************************************'

			loRango = loHoja.Range("B" & lnFilaActual)
			loRango.Value = "Artículo"
			
			loRango = loHoja.Range("C" & lnFilaActual)
			loRango.Value = "Modelo"
			
			loRango = loHoja.Range("D" & lnFilaActual)
			loRango.Value = "Descripción"
			
			loRango = loHoja.Range("E" & lnFilaActual)
			loRango.Value = "Stock" & vbLf & "Actual"
			
			loRango = loHoja.Range("F" & lnFilaActual)
			loRango.Value = "Stock" & vbLf & "Disponible"
			
			loRango = loHoja.Range("G" & lnFilaActual)
			loRango.Value = lcTituloCampoMonto & vbLf & "Unitario"
			
									
			loRango = loHoja.Range("B" & lnFilaActual & ":G" & lnFilaActual)
			loFuente = loRango.Font
			loFuente.Bold = True
			loFuente.Color = Rgb(255, 255, 255)
			loRango.Interior.Color = Rgb(0, 51, 153)
			
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
			loRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
			
			Dim lnFilaInicio As Integer  = lnFilaActual
			For lnRenglon As Integer = 0 To loDatos.Rows.Count - 1
				Dim loRenglon As DataRow = loDatos.Rows(lnRenglon)
				
				lnFilaActual += 1
			
				'# Factura
				loRango = loHoja.Range("B" & lnFilaActual)
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Cod_Art")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'Cod.Prov.
				loRango = loHoja.Range("C" & lnFilaActual)
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Modelo")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'Proveedor
				loRango = loHoja.Range("D" & lnFilaActual)
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Nom_Art")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'Fecha 
				loRango = loHoja.Range("E" & lnFilaActual)
				loRango.NumberFormat = lcFormatoCantidad
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Exi_Act1")), lnDecimalesCantidad)
								
				'Cantidad
				loRango = loHoja.Range("F" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoCantidad
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Exi_Dis1")), lnDecimalesCantidad)
					
				'Cantidad
				loRango = loHoja.Range("G" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoMontos
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Monto")), lnDecimalesMonto)				
				 
			Next lnRenglon
			
			Dim lnTotal As Integer = loDatos.Rows.Count
			loRango = loHoja.Range("B" & (lnFilaInicio) & ":G" & (lnFilaInicio))
			loRango.Select() 
			loExcel.Selection.AutoFilter()
			
			loRango = loHoja.Range("B" & (lnFilaInicio+1) & ":G" & (lnFilaInicio + lnTotal))
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
					
			Dim lnDesde AS Integer = lnFilaInicio
			Dim lnHasta AS Integer = lnFilaInicio + lnTotal
			
			lnFilaInicio += lnTotal + 2
			loRango = loHoja.Range("B" & (lnFilaInicio))
			'loRango.MergeCells = True
			loRango.NumberFormat = "@"
			loRango.Value = "Total Artículos: " & lnTotal.ToString()
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

			'loRango = loHoja.Range("H" & (lnFilaInicio))
			'loRango.NumberFormat = "@"
			'loRango.Value = "Total General: "
			'loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

			'loRango = loHoja.Range("I" & (lnFilaInicio))
			'loRango.NumberFormat = lcFormatoCantidad
			'loRango.Formula = "=SUM(I" & lnDesde & ":I" & lnHasta	& ")"

			'loRango = loHoja.Range("L" & (lnFilaInicio))
			'loRango.NumberFormat = lcFormatoMontos
			'loRango.Formula = "=SUM(L" & lnDesde & ":L" & lnHasta	& ")"

			'loRango = loHoja.Range("M" & (lnFilaInicio))
			'loRango.NumberFormat = lcFormatoMontos
			'loRango.Formula = "=SUM(M" & lnDesde & ":M" & lnHasta	& ")"

			'loRango = loHoja.Range("N" & (lnFilaInicio))
			'loRango.NumberFormat = lcFormatoMontos
			'loRango.Formula = "=SUM(N" & lnDesde & ":N" & lnHasta	& ")"

			loRango = loHoja.Range("B" & (lnFilaInicio) & ":G" & (lnFilaInicio))
			loFuente = loRango.Font
			loFuente.Bold = True
					
			loFilas = loCeldas.Rows
			loFilas.AutoFit()
			
			loColumnas = loCeldas.Rows
			loColumnas.AutoFit()
			
			loRango = loHoja.Range("B1:B" & lnFilaInicio)
			loRango.ColumnWidth = 35
			
			loRango = loHoja.Range("C1:C" & lnFilaInicio)
			loRango.ColumnWidth = 28
			
			loRango = loHoja.Range("D1:D" & lnFilaInicio)
			loRango.ColumnWidth = 70
			
			loRango = loHoja.Range("E1:E" & lnFilaInicio)
			loRango.ColumnWidth = 11
			
			loRango = loHoja.Range("F1:F" & lnFilaInicio)
			loRango.ColumnWidth = 13
			
			loRango = loHoja.Range("G1:G" & lnFilaInicio)
			loRango.ColumnWidth = 16
			
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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 23/01/13: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' RJG: 24/01/13: Se agregó el campo Tasa y se programo los cálculos correspondientes.		'
'-------------------------------------------------------------------------------------------'
