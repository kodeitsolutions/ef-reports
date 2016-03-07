'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCPagar_TiposProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class rCPagar_TiposProveedores
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
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
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT 	Cuentas_Pagar.Documento		AS Documento, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Cod_Tip		AS Cod_Tip, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Fec_Ini		AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Fec_Fin		AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Cod_Pro		AS Cod_Pro, ")
            loComandoSeleccionar.AppendLine("         	Proveedores.Nom_Pro			AS Nom_Pro, ")
            loComandoSeleccionar.AppendLine("         	Proveedores.Cod_Tip			AS Tip_Pro, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Cod_Ven		AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Cod_Tra		AS Cod_Tra, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Cod_Mon		AS Cod_Mon, ")
            loComandoSeleccionar.AppendLine("         	Cuentas_Pagar.Control		AS Control, ")
            If lcParametro11Desde.ToString = "Si" Then
                loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Comentario	AS Comentario, ")
            Else
                loComandoSeleccionar.AppendLine("			' '							AS Comentario, ")
            End If

            loComandoSeleccionar.AppendLine("			(CASE WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Bru *(-1) ELSE Cuentas_Pagar.Mon_Bru End)	AS Mon_Bru,  ")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Imp1																			AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Net *(-1) ELSE Cuentas_Pagar.Mon_Net End)	AS Mon_Net,  ")
            loComandoSeleccionar.AppendLine("			(CASE WHEN Tip_Doc = 'Credito' THEN Cuentas_Pagar.Mon_Sal *(-1) ELSE Cuentas_Pagar.Mon_Sal End)	AS Mon_Sal,  ")
            loComandoSeleccionar.AppendLine("			COALESCE(Cheques.Tip_Ope, '') AS Tip_Ope,")
            loComandoSeleccionar.AppendLine("			COALESCE(Cheques.Num_Doc, '') AS Num_Doc,")
            loComandoSeleccionar.AppendLine("			COALESCE(Cheques.Cod_Ban, '') AS Cod_Ban,")
            loComandoSeleccionar.AppendLine("			COALESCE(Cheques.Cod_Cue, '') AS Cod_Cue ")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine(" 	JOIN 	Proveedores	ON (Cuentas_Pagar.Cod_Pro = Proveedores.Cod_Pro) ")
            loComandoSeleccionar.AppendLine("	LEFT JOIN (")
            loComandoSeleccionar.AppendLine("				SELECT		Renglones_Pagos.Cod_Tip,")
            loComandoSeleccionar.AppendLine("							Renglones_Pagos.Doc_Ori,")
            loComandoSeleccionar.AppendLine("							Detalles_Pagos.Tip_Ope,")
            loComandoSeleccionar.AppendLine("							Detalles_Pagos.Num_Doc,")
            loComandoSeleccionar.AppendLine("							Detalles_Pagos.Cod_Ban,")
            loComandoSeleccionar.AppendLine("							Detalles_Pagos.Cod_Cue")
            loComandoSeleccionar.AppendLine("				FROM		Renglones_Pagos")
            loComandoSeleccionar.AppendLine("					JOIN	Pagos ")
            loComandoSeleccionar.AppendLine("						ON	Pagos.Documento = Renglones_Pagos.Documento")
            loComandoSeleccionar.AppendLine("						AND Pagos.Automatico = 1")
            loComandoSeleccionar.AppendLine("					JOIN	Detalles_Pagos ")
            loComandoSeleccionar.AppendLine("						ON	Pagos.Documento = Detalles_Pagos.Documento")
            loComandoSeleccionar.AppendLine("						AND	Detalles_Pagos.Num_Doc > ''")
            loComandoSeleccionar.AppendLine("				WHERE	Renglones_Pagos.Cod_Tip = 'ADEL'")
            loComandoSeleccionar.AppendLine("			) AS Cheques")
            loComandoSeleccionar.AppendLine("		ON	Cuentas_Pagar.Cod_Tip = Cheques.Cod_Tip")
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Documento = Cheques.Doc_Ori")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Documento BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Ven BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Status IN ( " & lcParametro4Desde & " ) ")
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Tra BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Tip BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("         	AND Proveedores.Cod_Tip BETWEEN " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("         	AND Cuentas_Pagar.Mon_Sal BETWEEN " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro13Hasta)


            If lcParametro10Desde = "Igual" Then
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev BETWEEN " & lcParametro9Desde)
            Else
                loComandoSeleccionar.AppendLine(" 				AND Cuentas_Pagar.Cod_Rev NOT BETWEEN " & lcParametro9Desde)
            End If
            loComandoSeleccionar.AppendLine("         AND " & lcParametro9Hasta)

            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCPagar_TiposProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCPagar_TiposProveedores.ReportSource = loObjetoReporte


            'Selección de opcion por excel (Microsoft Excel - xls)
            If (Me.Request.QueryString("salida").ToLower = "xls") Then
                ' Ruta donde se creara temporalmente el archivo
                Dim lcFileName As String = Server.MapPath("~\Administrativo\Temporales\rCPagar_TiposProveedores_" & Guid.NewGuid().ToString("N") & ".xls")
                ' Se exporta para crear el archivo temporal
                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, lcFileName)

                ' Se modifica el contenido del archivo
                Me.mGenerarArchivoExcel(lcFileName, laDatosReporte.Tables(0), "")

                ' Se coloca en la respuesta para decargar
                Me.Response.Clear()
                Me.Response.Buffer = True 
                Me.Response.AppendHeader("content-disposition", "attachment; filename=rCPagar_TiposProveedores.xls")
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

	Private Sub mGenerarArchivoExcel(ByVal lcNombreArchivo As String, ByVal loDatos As DataTable, ByVal lcParametrosReporte As String)
		
		Dim lnDecimales As Integer = goOpciones.pnDecimalesParaMonto
		Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
		
		Dim lcFormatoMontos As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimales)
		
		Dim lcFormatoCantidad As String 
		If (lnDecimalesCantidad > 0) Then 
			lcFormatoCantidad = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesCantidad)
		Else
			lcFormatoCantidad = "###,###,###,###,##0"
		End If
		

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

            loRango = loHoja.Range("B5:F5")
            loRango.Select()
            loRango.MergeCells = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.Value = "Listado de Cuentas por Pagar por Tipo de Proveedores"
            loFuente = loRango.Font
            loFuente.Size = 14
            loFuente.Bold = True

            ' Fecha y hora de creacion
			Dim ldFecha As DateTime = Date.Now()
			loRango = loHoja.Range("M1")
			loRango.NumberFormat = "@"
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
			loRango.Value = ldFecha.ToString("dd/MM/yyyy")
			loRango = loHoja.Range("M2")
			loRango.NumberFormat = "@"
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
			loRango.Value = ldFecha.ToString("hh:mm:ss tt")

            ' Parametros del reporte
            loRango = loHoja.Range("A7:M7")
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
			loRango.Value = "Tipo de"  & Strings.Chr(10) & "Proveedor"
			
			loRango = loHoja.Range("C" & lnFilaActual)
			loRango.Value = "Proveedor"
			
			loRango = loHoja.Range("D" & lnFilaActual)
			loRango.Value = "Nombre"
			
			loRango = loHoja.Range("E" & lnFilaActual)
			loRango.Value = "Tipo de "  & Strings.Chr(10) & "Documento"
			
			loRango = loHoja.Range("F" & lnFilaActual)
			loRango.Value = "Documento"
			
			loRango = loHoja.Range("G" & lnFilaActual)
			loRango.Value = "Emisión"
			
			loRango = loHoja.Range("H" & lnFilaActual)
			loRango.Value = "Vencimiento"
			
			loRango = loHoja.Range("I" & lnFilaActual)
			loRango.Value = "Num. de Cheque/"  & Strings.Chr(10) & "Transferencia"
			
			loRango = loHoja.Range("J" & lnFilaActual)
			loRango.Value = "Moneda"
			
			loRango = loHoja.Range("K" & lnFilaActual)
			loRango.Value = "Monto Bruto"
			
			loRango = loHoja.Range("L" & lnFilaActual)
			loRango.Value = "Monto"  & Strings.Chr(10) & "Impuesto"
			
			loRango = loHoja.Range("M" & lnFilaActual)
			loRango.Value = "Monto Neto"
						
			loRango = loHoja.Range("N" & lnFilaActual)
			loRango.Value = "Saldo"
						
			loRango = loHoja.Range("B" & lnFilaActual & ":N" & lnFilaActual)
			loFuente = loRango.Font
			loFuente.Bold = True
			
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
			loRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
			
			Dim lnFilaInicio As Integer  = lnFilaActual
			For lnRenglon As Integer = 0 To loDatos.Rows.Count - 1
				 Dim loRenglon As DataRow = loDatos.Rows(lnRenglon)
				 
				 lnFilaActual += 1
				 
				 'Tipo de Prov.
				 loRango = loCeldas(lnFilaActual, 2) 
				 loRango.NumberFormat = "@"
				 loRango.Value = CStr(loRenglon("Tip_Pro")).Trim()
				 loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				 
				 'Proveedor
				 loRango = loCeldas(lnFilaActual, 3) 
				 loRango.NumberFormat = "@"
				 loRango.Value = CStr(loRenglon("Cod_Pro")).Trim()
				 loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				 
				 'Nombre
				 loRango = loCeldas(lnFilaActual, 4) 
				 loRango.NumberFormat = "@"
				 loRango.Value = CStr(loRenglon("Nom_Pro")).Trim()
				 
				 'Tipo de Documento
				 loRango = loCeldas(lnFilaActual, 5) 
				 loRango.NumberFormat = "@"
				 loRango.Value = CStr(loRenglon("Cod_Tip")).Trim()
					
				 'Documento
				 loRango = loCeldas(lnFilaActual, 6) 
				 loRango.NumberFormat = "@"
				 loRango.Value = CStr(loRenglon("Documento")).Trim()
					
				 'Emisión
				 loRango = loCeldas(lnFilaActual, 7) 
				 loRango.NumberFormat = "@"
				 loRango.Value = CDate(loRenglon("Fec_Ini")).ToString("MM/dd/yyyy")
				 loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
				 
				 'Vencimiento
				 loRango = loCeldas(lnFilaActual, 8) 
				 loRango.NumberFormat = "@"
				 loRango.Value = CDate(loRenglon("Fec_Ini")).ToString("MM/dd/yyyy")
				 loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
				 
				 'Num. de Cheque/Transferencia
				 loRango = loCeldas(lnFilaActual, 9) 
				 loRango.NumberFormat = "@"
				 loRango.Value = CStr(loRenglon("Num_Doc")).Trim()

				 
				 ' Moneda
				 loRango = loCeldas(lnFilaActual, 10) 
				 loRango.NumberFormat = "@"
				 loRango.Value = CStr(loRenglon("Cod_Mon")).Trim()
				 
				 'Monto Bruto
				 loRango = loCeldas(lnFilaActual, 11) 
				 loRango.NumberFormat = lcFormatoMontos
				 loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Mon_Bru")), lnDecimales)
				 
				 'Monto Impuesto
				 loRango = loCeldas(lnFilaActual, 12) 
				 loRango.NumberFormat = lcFormatoMontos
				 loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Mon_Imp1")), lnDecimales)
				 
				 'Monto Neto
				 loRango = loCeldas(lnFilaActual, 13) 
				 loRango.NumberFormat = lcFormatoMontos
				 loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Mon_Net")), lnDecimales)
				 
				 'Monto Saldo
				 loRango = loCeldas(lnFilaActual, 14) 
				 loRango.NumberFormat = lcFormatoMontos
				 loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Mon_Sal")), lnDecimales)
				 
			Next lnRenglon
			
			Dim lnTotal As Integer = loDatos.Rows.Count
			loRango = loHoja.Range("B" & (lnFilaInicio+1) & ":N" & (lnFilaInicio + lnTotal))
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
					
			loRango = loHoja.Range("B" & (lnFilaInicio+1) & ":N" & (lnFilaInicio + lnTotal))
			
			Dim lnDesde AS Integer = lnFilaInicio
			Dim lnHasta AS Integer = lnFilaInicio + lnTotal
			
			lnFilaInicio += lnTotal + 2
			loRango = loHoja.Range("B" & (lnFilaInicio))
			loRango.NumberFormat = "@"
			loRango.Value = "Total Documentos: " & lnTotal.ToString()
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

			loRango = loHoja.Range("J" & (lnFilaInicio))
			loRango.NumberFormat = "@"
			loRango.Value = "Total General: "
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

			loRango = loHoja.Range("K" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoMontos
			loRango.Value = "=SUM(K" & lnDesde & ":K" & lnHasta	& ")"

			loRango = loHoja.Range("L" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoMontos
			loRango.Value = "=SUM(L" & lnDesde & ":L" & lnHasta	& ")"

			loRango = loHoja.Range("M" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoCantidad
			loRango.Value = "=SUM(M" & lnDesde & ":M" & lnHasta	& ")"

			loRango = loHoja.Range("N" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoCantidad
			loRango.Value = "=SUM(N" & lnDesde & ":N" & lnHasta	& ")"

			loRango = loHoja.Range("B" & (lnFilaInicio) & ":N" & (lnFilaInicio))
			loFuente = loRango.Font
			loFuente.Bold = True
					
			loFilas = loCeldas.Rows
			loFilas.AutoFit()
			
			loColumnas = loCeldas.Rows
			loColumnas.AutoFit()
			
			loRango = loHoja.Range("B1:B" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("C1:C" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("D1:D" & lnFilaInicio)
			loRango.ColumnWidth = 60
			
			loRango = loHoja.Range("E1:E" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("F1:F" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("G1:G" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("H1:H" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("I1:I" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("J1:J" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("K1:K" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("L1:L" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("M1:M" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("N1:N" & lnFilaInicio)
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
' RJG: 25/08/12: Programacion inicial (A partir de rCPagar_Proveedores_2).					'
'-------------------------------------------------------------------------------------------'
