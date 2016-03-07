'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEfectividadVendedores_CotizadosVsFacturados"
'-------------------------------------------------------------------------------------------'
Partial Class rEfectividadVendedores_CotizadosVsFacturados
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
            'Dim lcParametro9Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12))
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12))
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13))
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13))
            Dim lcParametro14Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(14))
            Dim lcParametro14Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(14))



            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()



            loComandoSeleccionar.AppendLine("SELECT		")
            loComandoSeleccionar.AppendLine("			Vendedores.Cod_Ven						AS Cod_Ven,")
            loComandoSeleccionar.AppendLine("			Vendedores.Nom_Ven						AS Nom_Ven,")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Cotizaciones.Can_Art1)	AS Can_Art1,")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Cotizaciones.Mon_Net)		AS Mon_Net")
            loComandoSeleccionar.AppendLine("INTO		#tmpCOTIZACIONES					")
            loComandoSeleccionar.AppendLine("FROM		Cotizaciones						")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Cotizaciones ON Renglones_Cotizaciones.Documento = Cotizaciones.Documento")
            loComandoSeleccionar.AppendLine("	JOIN	Articulos ON Renglones_Cotizaciones.Cod_Art = Articulos.Cod_Art")
            loComandoSeleccionar.AppendLine("	JOIN	Almacenes ON Renglones_Cotizaciones.Cod_Alm = Almacenes.Cod_Alm")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes ON Clientes.Cod_Cli = Cotizaciones.Cod_CLi")
            loComandoSeleccionar.AppendLine("	JOIN	Vendedores ON Vendedores.Cod_Ven = Cotizaciones.Cod_Ven")

            loComandoSeleccionar.AppendLine("WHERE      ")
            loComandoSeleccionar.AppendLine(" 			Cotizaciones.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Cotizaciones.Cod_Art BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_tip BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Renglones_Cotizaciones.Cod_Alm BETWEEN " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro8Hasta)
            'loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Status IN (" & lcParametro9Desde & ")")
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Status IN ('Confirmado', 'Afectado', 'Procesado')")
            loComandoSeleccionar.AppendLine(" 			AND Articulos.cod_Mar BETWEEN " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Zon BETWEEN " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Cla BETWEEN " & lcParametro11Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro11Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Tip BETWEEN " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Suc BETWEEN " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine(" 			AND Cotizaciones.Cod_Rev BETWEEN " & lcParametro14Desde)
            loComandoSeleccionar.AppendLine(" 			AND " & lcParametro14Hasta)
            loComandoSeleccionar.AppendLine("GROUP BY	Vendedores.Cod_Ven, Vendedores.Nom_Ven")
																		  

            loComandoSeleccionar.AppendLine("SELECT		")
            loComandoSeleccionar.AppendLine("        	Facturas.Cod_Ven												AS Cod_Ven,	")
            loComandoSeleccionar.AppendLine("			SUM(Renglones_Facturas.Can_Art1)								AS Can_Art1,")
            loComandoSeleccionar.AppendLine("        	SUM(Renglones_Facturas.Mon_Net)									AS Mon_Net,	")
            loComandoSeleccionar.AppendLine("        	SUM(Renglones_Facturas.Cos_Ult1*Renglones_Facturas.Can_Art1)	AS Cos_Fac	")
            loComandoSeleccionar.AppendLine("INTO		#tmpFACTURAS	")
            loComandoSeleccionar.AppendLine("FROM		Facturas		")
            loComandoSeleccionar.AppendLine("   JOIN	Renglones_Facturas ON Renglones_Facturas.Documento = Facturas.Documento      ")
            loComandoSeleccionar.AppendLine("   JOIN	Articulos ON Renglones_Facturas.Cod_Art = Articulos.Cod_Art      ")
            loComandoSeleccionar.AppendLine("   JOIN	Almacenes ON Renglones_Facturas.Cod_Alm = Almacenes.Cod_Alm      ")
            loComandoSeleccionar.AppendLine("   JOIN	Clientes ON Clientes.Cod_Cli = Facturas.Cod_CLi			 ")
            loComandoSeleccionar.AppendLine("   JOIN	Vendedores ON Vendedores.Cod_Ven = Facturas.Cod_Ven      ")

			loComandoSeleccionar.AppendLine("WHERE      ")
			loComandoSeleccionar.AppendLine(" 			Facturas.Fec_Ini BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Cli BETWEEN " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Ven BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Renglones_Facturas.Cod_Art BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Dep BETWEEN " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro4Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Sec BETWEEN " & lcParametro5Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro5Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Tip BETWEEN " & lcParametro6Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro6Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Cla BETWEEN " & lcParametro7Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro7Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Renglones_Facturas.Cod_Alm BETWEEN " & lcParametro8Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro8Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Facturas.Status IN ('Confirmado', 'Afectado', 'Procesado')")
			loComandoSeleccionar.AppendLine(" 			AND Articulos.Cod_Mar BETWEEN " & lcParametro9Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro9Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Zon BETWEEN " & lcParametro10Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro10Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Cla BETWEEN " & lcParametro11Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro11Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Clientes.Cod_Tip BETWEEN " & lcParametro12Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro12Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Suc BETWEEN " & lcParametro13Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro13Hasta)
			loComandoSeleccionar.AppendLine(" 			AND Facturas.Cod_Rev BETWEEN " & lcParametro14Desde)
			loComandoSeleccionar.AppendLine(" 			AND " & lcParametro14Hasta)
			loComandoSeleccionar.AppendLine("GROUP BY	Facturas.Cod_Ven")
			loComandoSeleccionar.AppendLine("")

			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("SELECT #tmpCOTIZACIONES.Cod_Ven 														AS Cod_Ven,  	")
			loComandoSeleccionar.AppendLine("   	#tmpCOTIZACIONES.Nom_Ven 														AS Nom_Ven,  	")
			loComandoSeleccionar.AppendLine("   	SUM(#tmpCOTIZACIONES.Can_Art1)													AS Can_Art1_Cot,")
			loComandoSeleccionar.AppendLine("   	SUM(COALESCE(#tmpFACTURAS.Can_Art1, 0))											AS Can_Art1_Fac,")
			loComandoSeleccionar.AppendLine("		CASE WHEN SUM(#tmpCOTIZACIONES.Can_Art1) > 0													")
			loComandoSeleccionar.AppendLine("			THEN SUM(COALESCE(#tmpFACTURAS.Can_Art1, 0))/ SUM(#tmpCOTIZACIONES.Can_Art1)*100			")
			loComandoSeleccionar.AppendLine("			ELSE 100																					")
			loComandoSeleccionar.AppendLine("		END																				AS Efic1,		")
			loComandoSeleccionar.AppendLine("		SUM(#tmpCOTIZACIONES.Mon_Net)													AS Mon_Net_Cot, ")
			loComandoSeleccionar.AppendLine("		SUM(COALESCE(#tmpFACTURAS.Mon_Net,0))											AS Mon_Net_Fac, ")
			loComandoSeleccionar.AppendLine("		CASE WHEN SUM(#tmpCOTIZACIONES.Mon_Net) > 0														")
			loComandoSeleccionar.AppendLine("			THEN SUM(COALESCE(#tmpFACTURAS.Mon_Net, 0)) / SUM(#tmpCOTIZACIONES.Mon_Net) * 100			")
			loComandoSeleccionar.AppendLine("			ELSE 100																					")
			loComandoSeleccionar.AppendLine("		END																				AS Efic2,		")
			loComandoSeleccionar.AppendLine("		SUM(COALESCE(#tmpFACTURAS.Cos_Fac,0))											AS Cos_Fac,		")
			loComandoSeleccionar.AppendLine("		CASE WHEN SUM(COALESCE(#tmpFACTURAS.Cos_Fac,0)) > 0												")
			loComandoSeleccionar.AppendLine("			THEN (SUM(COALESCE(#tmpFACTURAS.Mon_Net, 0)) - SUM(COALESCE(#tmpFACTURAS.Cos_Fac,0)))/SUM(COALESCE(#tmpFACTURAS.Cos_Fac,0)) * 100	")
			loComandoSeleccionar.AppendLine("			ELSE 100																					")
			loComandoSeleccionar.AppendLine("		END																				AS Utilidad		")
			loComandoSeleccionar.AppendLine("FROM	#tmpCOTIZACIONES")
			loComandoSeleccionar.AppendLine("	LEFT JOIN #tmpFACTURAS ON #tmpFACTURAS.Cod_Ven = #tmpCOTIZACIONES.Cod_Ven")
			loComandoSeleccionar.AppendLine("GROUP BY	#tmpCOTIZACIONES.Cod_Ven, #tmpCOTIZACIONES.Nom_Ven")
			loComandoSeleccionar.AppendLine("ORDER BY      " & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpCOTIZACIONES")
			loComandoSeleccionar.AppendLine("DROP TABLE #tmpFACTURAS")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())


			Dim loServicios As New cusDatos.goDatos
			Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEfectividadVendedores_CotizadosVsFacturados", laDatosReporte)

			Me.mTraducirReporte(loObjetoReporte)
			Me.mFormatearCamposReporte(loObjetoReporte)
			Me.crvrEfectividadVendedores_CotizadosVsFacturados.ReportSource = loObjetoReporte


            'Selección de opcion por excel (Microsoft Excel - xls)
            If (Me.Request.QueryString("salida").ToLower = "xls") Then
                ' Ruta donde se creara temporalmente el archivo
                Dim lcFileName As String = Server.MapPath("~\Administrativo\Temporales\rEfectividadVendedores_CotizadosVsFacturados_" & Guid.NewGuid().ToString("N") & ".xls")
                ' Se exporta para crear el archivo temporal
                loObjetoReporte.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, lcFileName)

                ' Se modifica el contenido del archivo
                Me.mGenerarArchivoExcel(lcFileName, laDatosReporte.Tables(0), "")

                ' Se coloca en la respuesta para decargar
                Me.Response.Clear()
                Me.Response.Buffer = True 
                Me.Response.AppendHeader("content-disposition", "attachment; filename=rEfectividadVendedores_CotizadosVsFacturados.xls")
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

            loRango = loHoja.Range("B5:J5")
            loRango.Select()
            loRango.MergeCells = True
            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            loRango.Value = "Efectividad Cotizaciones Vs Facturaciones por Vendedor"
            loFuente = loRango.Font
            loFuente.Size = 14
            loFuente.Bold = True

            ' Fecha y hora de creacion
			Dim ldFecha As DateTime = Date.Now()
			loRango = loHoja.Range("J1")
			loRango.NumberFormat = "@"
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
			loRango.Value = ldFecha.ToString("dd/MM/yyyy")
			
			loRango = loHoja.Range("J2")
			loRango.NumberFormat = "@" 'La celda almacena un string
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
			loRango.Value = ldFecha.ToString("hh:mm:ss tt")

            ' Parametros del reporte
            loRango = loHoja.Range("B7:J7")
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
			loRango.Value = "Código"
			
			loRango = loHoja.Range("C" & lnFilaActual)
			loRango.Value = "Vendedor"
			
			loRango = loHoja.Range("D" & lnFilaActual)
			loRango.Value = "Cotizaciones"
			
			loRango = loHoja.Range("E" & lnFilaActual)
			loRango.Value = "Facturaciones"
			
			loRango = loHoja.Range("F" & lnFilaActual)
			loRango.Value = "%"
			
			loRango = loHoja.Range("G" & lnFilaActual)
			loRango.Value = "Monto Neto" & vbLf & "Cotizaciones"
			
			loRango = loHoja.Range("H" & lnFilaActual)
			loRango.Value = "Monto Neto" & vbLf & "Facturaciones"
			
			loRango = loHoja.Range("I" & lnFilaActual)
			loRango.Value = "%"
			
			loRango = loHoja.Range("J" & lnFilaActual)
			loRango.Value = "Margen de" & vbLf & "Utilidad"
									
			loRango = loHoja.Range("B" & lnFilaActual & ":J" & lnFilaActual)
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
			
				'Código
				loRango = loHoja.Range("B" & lnFilaActual)
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Cod_Ven")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'Vendedor
				loRango = loHoja.Range("C" & lnFilaActual)
				loRango.NumberFormat = "@"
				loRango.Value = CStr(loRenglon("Nom_Ven")).Trim()
				loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
				
				'Cantidad Cotizado
				loRango = loHoja.Range("D" & lnFilaActual)
				loRango.NumberFormat = lcFormatoCantidad
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Can_Art1_Cot")), lnDecimalesCantidad)
				
				'Cantidad Facturado 
				loRango = loHoja.Range("E" & lnFilaActual)
				loRango.NumberFormat = lcFormatoCantidad
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Can_Art1_Fac")), lnDecimalesCantidad)
				
				'% Eficiencia 1 
				loRango = loHoja.Range("F" & lnFilaActual)	
				loRango.NumberFormat = lcFormatoPorcentaje
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Efic1")), lnDecimalesPorcentaje)
						
				'Neto Cotizado
				loRango = loHoja.Range("G" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoMontos '#.###.##0,00	
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Mon_Net_Cot")), lnDecimalesMonto)
					
				'Neto Facturado
				loRango = loHoja.Range("H" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoMontos
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Mon_Net_Fac")), lnDecimalesMonto)
				
				'% Eficiencia 2
				loRango = loHoja.Range("I" & lnFilaActual) 
				loRango.NumberFormat = lcFormatoPorcentaje	
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Efic2")), lnDecimalesPorcentaje)
					
				'%	Margen
				loRango = loHoja.Range("J" & lnFilaActual)   
				loRango.NumberFormat = lcFormatoPorcentaje
				loRango.Value = goServicios.mRedondearValor(CDec(loRenglon("Utilidad")), lnDecimalesPorcentaje)
				 
			Next lnRenglon
			
			Dim lnTotal As Integer = loDatos.Rows.Count
			loRango = loHoja.Range("B" & (lnFilaInicio) & ":J" & (lnFilaInicio))
			loRango.Select() 
			loExcel.Selection.AutoFilter()
			
			loRango = loHoja.Range("B" & (lnFilaInicio+1) & ":J" & (lnFilaInicio + lnTotal))
			loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)
					
			Dim lnDesde AS Integer = lnFilaInicio
			Dim lnHasta AS Integer = lnFilaInicio + lnTotal
			
			lnFilaInicio += lnTotal + 2
			loRango = loHoja.Range("B" & (lnFilaInicio))
			'loRango.MergeCells = True
			loRango.NumberFormat = "@"
			loRango.Value = "Total Vendedores: " & lnTotal.ToString()
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

			loRango = loHoja.Range("C" & (lnFilaInicio))
			loRango.NumberFormat = "@"
			loRango.Value = "Total General: "
			loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

			loRango = loHoja.Range("D" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoMontos
			loRango.Formula = "=SUM(D" & lnDesde & ":D" & lnHasta	& ")"

			loRango = loHoja.Range("E" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoMontos
			loRango.Formula = "=SUM(E" & lnDesde & ":E" & lnHasta	& ")"

			loRango = loHoja.Range("F" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoPorcentaje
			loRango.Formula = "=IF(D" & (lnFilaInicio) & ">0, E" & (lnFilaInicio) & "*100/D" & (lnFilaInicio) & ", 100)"
			
			loRango = loHoja.Range("G" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoMontos
			loRango.Formula = "=SUM(G" & lnDesde & ":G" & lnHasta	& ")"

			loRango = loHoja.Range("H" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoMontos
			loRango.Formula = "=SUM(H" & lnDesde & ":H" & lnHasta	& ")"

			loRango = loHoja.Range("I" & (lnFilaInicio))
			loRango.NumberFormat = lcFormatoPorcentaje
			loRango.Formula = "=IF(G" & (lnFilaInicio) & ">0, H" & (lnFilaInicio) & "*100/G" & (lnFilaInicio) & ", 100)"
			
			'loRango = loHoja.Range("J" & (lnFilaInicio))
			'loRango.NumberFormat = lcFormatoMontos
			'loRango.Formula = "=IF(G" & (lnFilaInicio) & ">0, I" & (lnFilaInicio) & "*100/G" & (lnFilaInicio) & ", 100)"

			loRango = loHoja.Range("B" & (lnFilaInicio) & ":J" & (lnFilaInicio))
			loFuente = loRango.Font
			loFuente.Bold = True
					
			loFilas = loCeldas.Rows
			loFilas.AutoFit()
			
			loColumnas = loCeldas.Rows
			loColumnas.AutoFit()
			
			loRango = loHoja.Range("B1:B" & lnFilaInicio)
			loRango.ColumnWidth = 18
			
			loRango = loHoja.Range("C1:C" & lnFilaInicio)
			loRango.ColumnWidth = 45
			
			loRango = loHoja.Range("D1:D" & lnFilaInicio)
			loRango.ColumnWidth = 12
			
			loRango = loHoja.Range("E1:E" & lnFilaInicio)
			loRango.ColumnWidth = 14
			
			loRango = loHoja.Range("F1:F" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("G1:G" & lnFilaInicio)
			loRango.ColumnWidth = 14
			
			loRango = loHoja.Range("H1:H" & lnFilaInicio)
			loRango.ColumnWidth = 14
			
			loRango = loHoja.Range("I1:I" & lnFilaInicio)
			loRango.ColumnWidth = 10
			
			loRango = loHoja.Range("J1:J" & lnFilaInicio)
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
' RJG: 09/10/12: Codigo inicial, a partir de r EfectividadArticulos_CotizadosVsFacturados.	'
'-------------------------------------------------------------------------------------------'
