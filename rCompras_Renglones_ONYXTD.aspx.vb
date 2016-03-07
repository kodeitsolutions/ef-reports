'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCompras_Renglones_ONYXTD"
'-------------------------------------------------------------------------------------------'
Partial Class rCompras_Renglones_ONYXTD
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
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT			Compras.Documento, ")
            loComandoSeleccionar.AppendLine("				Compras.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("				Compras.Fec_Fin, ")
            loComandoSeleccionar.AppendLine("				Compras.Cod_Pro, ")
            loComandoSeleccionar.AppendLine("				Compras.Cod_Tra, ")
            loComandoSeleccionar.AppendLine("				Compras.Cod_For, ")
            loComandoSeleccionar.AppendLine("				Compras.Cod_Ven, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Cod_Art, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Renglon, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Cod_Alm, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Precio1, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Can_Art1, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Cod_Uni, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Por_Des, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Mon_Bru, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Mon_Des, ")
            loComandoSeleccionar.AppendLine("				Renglones_Compras.Mon_Net, ")
            loComandoSeleccionar.AppendLine("				Articulos.Nom_Art, ")
            loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro, ")
            loComandoSeleccionar.AppendLine("				Transportes.Nom_Tra, ")
            loComandoSeleccionar.AppendLine("				Formas_Pagos.Nom_For, ")
            loComandoSeleccionar.AppendLine("				Articulos.Cod_Mar ")
            loComandoSeleccionar.AppendLine("FROM			Compras")
            loComandoSeleccionar.AppendLine("	JOIN		Renglones_Compras")
            loComandoSeleccionar.AppendLine("		ON		Compras.Documento = Renglones_Compras.Documento")
            loComandoSeleccionar.AppendLine("	JOIN		Articulos")
            loComandoSeleccionar.AppendLine("		ON		Renglones_Compras.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN		Proveedores")
            loComandoSeleccionar.AppendLine("		ON		Compras.Cod_Pro = Proveedores.Cod_Pro")
            loComandoSeleccionar.AppendLine("	JOIN		Formas_Pagos")
            loComandoSeleccionar.AppendLine("		ON		Compras.Cod_For = Formas_Pagos.Cod_For")
            loComandoSeleccionar.AppendLine("	JOIN		Transportes ")
            loComandoSeleccionar.AppendLine("		ON		Compras.Cod_Tra = Transportes.Cod_Tra")
            loComandoSeleccionar.AppendLine(" WHERE			Compras.Documento BETWEEN ''")
            loComandoSeleccionar.AppendLine("	 				AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Compras.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 					AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Compras.Cod_Pro BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 					AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Compras.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine(" 				AND Compras.Cod_Rev BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 					AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 				AND Compras.Cod_Suc BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 					AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos
			
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCompras_Renglones_ONYXTD", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCompras_Renglones_ONYXTD.ReportSource = loObjetoReporte

            'Selección de opcion por excel (Microsoft Excel - xls)
            If (Me.Request.QueryString("salida").ToLower() = "xls") Then
                ' Ruta donde se creara temporalmente el archivo
                Dim lcFileName As String = Server.MapPath("~\Administrativo\Temporales\rCompras_Renglones_ONYXTD_" & Guid.NewGuid().ToString("N") & ".xls")
                ' Se exporta para crear el archivo temporal
                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, lcFileName)

                ' Se modifica el contenido del archivo
                Dim lcParametrosReporte As String = cusAplicacion.goReportes.mObtenerParametros(cusAplicacion.goReportes.paNombresParametros, cusAplicacion.goReportes.paParametrosIniciales, cusAplicacion.goReportes.paParametrosFinales)
                Me.mGenerarArchivoExcel(lcFileName, laDatosReporte.Tables(0), lcParametrosReporte)

                ' Se coloca en la respuesta para decargar
                Me.Response.Clear()
                Me.Response.Buffer = True 
                Me.Response.AppendHeader("content-disposition", "attachment; filename=rCompras_Renglones_ONYXTD.xls")
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

            loRango = loHoja.Range("B5:O5")
            loRango.Select()
            loRango.MergeCells = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.Value = "Listado de Compras con sus Renglones (ONYXTD)"
            loFuente = loRango.Font
            loFuente.Size = 14
            loFuente.Bold = True

            ' Fecha y hora de creacion
			Dim ldFecha As DateTime = Date.Now()
			loRango = loHoja.Range("O1")
			loRango.NumberFormat = "mm/dd/yyyy;@"
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
			loRango.Value = ldFecha
			
			loRango = loHoja.Range("O2")
			loRango.NumberFormat = "@" 'La celda almacena un string
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
			loRango.Value = ldFecha.ToString("hh:mm:ss tt")

            ' Parametros del reporte
            loRango = loHoja.Range("B7:O7")
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
			loRango.Value = "# Factura"
			
			loRango = loHoja.Range("C" & lnFilaActual)
			loRango.Value = "Cod.Prov."
			
			loRango = loHoja.Range("D" & lnFilaActual)
			loRango.Value = "Proveedor"
			
			loRango = loHoja.Range("E" & lnFilaActual)
			loRango.Value = "Fecha"
			
			loRango = loHoja.Range("F" & lnFilaActual)
			loRango.Value = "Comprador"
			
			loRango = loHoja.Range("G" & lnFilaActual)
			loRango.Value = "Código de Producto"
			
			loRango = loHoja.Range("H" & lnFilaActual)
			loRango.Value = "Producto"
			
			loRango = loHoja.Range("I" & lnFilaActual)
			loRango.Value = "Cantidad"
			
			loRango = loHoja.Range("J" & lnFilaActual)
			loRango.Value = "Tipo" & vbLf & "Unidad"
									
			loRango = loHoja.Range("K" & lnFilaActual)
			loRango.Value = "Costo"
									
			loRango = loHoja.Range("L" & lnFilaActual)
			loRango.Value = "Monto" & vbLf & "Bruto"
									
			loRango = loHoja.Range("M" & lnFilaActual)
			loRango.Value = "Impuesto"
									
			loRango = loHoja.Range("N" & lnFilaActual)
			loRango.Value = "Monto" & vbLf & "Neto"
									
			loRango = loHoja.Range("O" & lnFilaActual)
			loRango.Value = "Marca"
									
			loRango = loHoja.Range("B" & lnFilaActual & ":O" & lnFilaActual)
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
				loRango.Value = CStr(loRenglon("Documento")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'Cod.Prov.
				loRango = loHoja.Range("C" & lnFilaActual)
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Cod_Pro")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'Proveedor
				loRango = loHoja.Range("D" & lnFilaActual)
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Nom_pro")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'Fecha 
				loRango = loHoja.Range("E" & lnFilaActual)
				loRango.NumberFormat = "mm/dd/yyyy;@"
				loRango.Value = CDate(loRenglon("Fec_Ini"))'CDate(loRenglon("Fec_Ini")).ToString("dd/MM/yyyy")
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
				
				'Comprador
				loRango = loHoja.Range("F" & lnFilaActual)	
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Cod_Ven")).Trim()
						
				'Cod. Producto
				loRango = loHoja.Range("G" & lnFilaActual) 
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Cod_Art")).Trim()
					
				'Producto
				loRango = loHoja.Range("H" & lnFilaActual) 
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Nom_Art")).Trim()
				
				'Cantidad
				loRango = loHoja.Range("I" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoPorcentaje	
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Can_Art1")), lnDecimalesMonto)
					
				'Cantidad
				loRango = loHoja.Range("I" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoCantidad
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Can_Art1")), lnDecimalesCantidad)
					
				'Tipo
				loRango = loHoja.Range("J" & lnFilaActual)   
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Cod_Uni")).Trim()
				
				'Costo
				loRango = loHoja.Range("K" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoMontos
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Precio1")), lnDecimalesMonto)
					
				'Monto Bruto
				loRango = loHoja.Range("L" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoMontos
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Mon_Net")), lnDecimalesMonto)
					
				'Impuesto
				loRango = loHoja.Range("M" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoMontos
				If (CDec(loRenglon("Mon_Imp1"))<>0D) Then
					loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Mon_Imp1")), lnDecimalesMonto)
					loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
				Else
					loRango.Value = "-"
					loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
				End If
					
				'Monto  Neto
				loRango = loHoja.Range("N" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoMontos
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Mon_Net")) + CDec(loRenglon("Mon_Imp1")), lnDecimalesMonto)
					
				'Marca
				loRango = loHoja.Range("O" & lnFilaActual) 
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Cod_Mar")).Trim()
				 
			Next lnRenglon
			
			Dim lnTotal As Integer = loDatos.Rows.Count
			loRango = loHoja.Range("B" & (lnFilaInicio) & ":O" & (lnFilaInicio))
			loRango.Select() 
			loExcel.Selection.AutoFilter()
			
			loRango = loHoja.Range("B" & (lnFilaInicio+1) & ":O" & (lnFilaInicio + lnTotal))
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
					
			Dim lnDesde AS Integer = lnFilaInicio
			Dim lnHasta AS Integer = lnFilaInicio + lnTotal
			
			lnFilaInicio += lnTotal + 2
			loRango = loHoja.Range("B" & (lnFilaInicio))
			'loRango.MergeCells = True
			loRango.NumberFormat = "@"
			loRango.Value = "Total Documentos: " & lnTotal.ToString()
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

			loRango = loHoja.Range("H" & (lnFilaInicio))
			loRango.NumberFormat = "@"
			loRango.Value = "Total General: "
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

			loRango = loHoja.Range("I" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoCantidad
			loRango.Formula = "=SUM(I" & lnDesde & ":I" & lnHasta	& ")"

			loRango = loHoja.Range("L" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoMontos
			loRango.Formula = "=SUM(L" & lnDesde & ":L" & lnHasta	& ")"

			loRango = loHoja.Range("M" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoMontos
			loRango.Formula = "=SUM(M" & lnDesde & ":M" & lnHasta	& ")"

			loRango = loHoja.Range("N" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoMontos
			loRango.Formula = "=SUM(N" & lnDesde & ":N" & lnHasta	& ")"

			loRango = loHoja.Range("B" & (lnFilaInicio) & ":O" & (lnFilaInicio))
			loFuente = loRango.Font
			loFuente.Bold = True
					
			loFilas = loCeldas.Rows
			loFilas.AutoFit()
			
			loColumnas = loCeldas.Rows
			loColumnas.AutoFit()
			
			loRango = loHoja.Range("B1:B" & lnFilaInicio)
			loRango.ColumnWidth = 13
			
			loRango = loHoja.Range("C1:C" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("D1:D" & lnFilaInicio)
			loRango.ColumnWidth = 40
			
			loRango = loHoja.Range("E1:E" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("F1:F" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("G1:G" & lnFilaInicio)
			loRango.ColumnWidth = 30
			
			loRango = loHoja.Range("H1:H" & lnFilaInicio)
			loRango.ColumnWidth = 50
			
			loRango = loHoja.Range("I1:I" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("J1:J" & lnFilaInicio)
			loRango.ColumnWidth = 9
			
			loRango = loHoja.Range("K1:K" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("L1:L" & lnFilaInicio)
			loRango.ColumnWidth = 14
			
			loRango = loHoja.Range("M1:M" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("N1:N" & lnFilaInicio)
			loRango.ColumnWidth = 14
			
			loRango = loHoja.Range("O1:O" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
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
' RJG:  23/01/13: Código inicial, a partir de rCompras_Renglones_ONYXTD.					'
'-------------------------------------------------------------------------------------------'
