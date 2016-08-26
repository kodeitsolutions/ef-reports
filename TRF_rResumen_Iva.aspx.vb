'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rResumen_Iva"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rResumen_Iva
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcAño As String = CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Year)
            Dim lcMes As String = CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Month)
            If lcMes < 10 Then
                lcMes = "0" + lcMes
            End If

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecIni			AS DATETIME")
            loComandoSeleccionar.AppendLine("DECLARE	@sp_FecFin			AS DATETIME")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SET	@sp_FecIni = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("SET	@sp_FecFin = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT SUM(Mon_Exe) AS TotalExentoC")
            loComandoSeleccionar.AppendLine("INTO #tmpExentoC")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("WHERE Mon_Exe <> 0")
            loComandoSeleccionar.AppendLine("	AND Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("	AND Fec_Reg BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT SUM(Mon_Bas1) AS TotalBaseC,")
            loComandoSeleccionar.AppendLine("		SUM(Mon_Imp1) AS TotalImpuestoC")
            loComandoSeleccionar.AppendLine("INTO #tmpBaseC")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("WHERE Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("	AND Fec_Reg BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT SUM(Retenciones_Documentos.Mon_Ret) AS TotalRetenidoC")
            loComandoSeleccionar.AppendLine("INTO #tmpRetenidoC")
            loComandoSeleccionar.AppendLine("FROM Retenciones_Documentos")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("		ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT ISNULL(SUM(Mon_Exe),0) AS TotalExentoV")
            loComandoSeleccionar.AppendLine("INTO #tmpExentoV")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("WHERE Mon_Exe <> 0")
            loComandoSeleccionar.AppendLine("	AND Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("	AND Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Renglones_Facturas.Documento AS Documento,")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Facturas.Mon_Net) AS TotalVentaNoGravada")
            loComandoSeleccionar.AppendLine("INTO #tmpVentaNoGravada")
            loComandoSeleccionar.AppendLine("FROM Renglones_Facturas")
            loComandoSeleccionar.AppendLine("	JOIN Facturas ON Facturas.Documento = Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE Articulos.Cod_Cla = 'VING'")
            loComandoSeleccionar.AppendLine("	AND Facturas.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND Facturas.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("GROUP BY Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT SUM(Mon_Bas1) AS BaseVFactura,")
            loComandoSeleccionar.AppendLine("		SUM(Mon_Imp1) AS ImpuestoVFactura")
            loComandoSeleccionar.AppendLine("INTO #tmpBaseVFactura")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("WHERE Cod_Tip = 'FACT'")
            loComandoSeleccionar.AppendLine("	AND Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT SUM(Renglones_Facturas.Mon_Bas1) AS BaseVFactura,")
            loComandoSeleccionar.AppendLine("		SUM(Renglones_Facturas.Mon_Imp1) AS ImpuestoVFactura")
            loComandoSeleccionar.AppendLine("FROM Renglones_Facturas")
            loComandoSeleccionar.AppendLine("	JOIN Facturas ON Facturas.Documento = Renglones_Facturas.Documento")
            loComandoSeleccionar.AppendLine("	JOIN Articulos ON Articulos.Cod_Art = Renglones_Facturas.Cod_Art")
            loComandoSeleccionar.AppendLine("WHERE Articulos.Cod_Cla <> 'VING'")
            loComandoSeleccionar.AppendLine("	AND Facturas.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND Facturas.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT ISNULL(SUM(Mon_Bru)*-1,0) AS BaseVNotaCre,")
            loComandoSeleccionar.AppendLine("		ISNULL(SUM(Mon_Imp1)*-1,0) AS ImpuestoVNotaCre")
            loComandoSeleccionar.AppendLine("INTO #tmpBaseVNotaCre")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("WHERE Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("	AND Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT ISNULL(SUM(Cuentas_Cobrar.Mon_Bru),0) AS TotalRetenidoV")
            loComandoSeleccionar.AppendLine("INTO #tmpRetenidoV")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Cobrar ")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Cobrar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Cobrar.Fec_Ini BETWEEN @sp_FecIni AND @sp_FecFin")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT #tmpExentoC.TotalExentoC,")
            loComandoSeleccionar.AppendLine("		#tmpBaseC.TotalBaseC,")
            loComandoSeleccionar.AppendLine("		#tmpBaseC.TotalImpuestoC,")
            loComandoSeleccionar.AppendLine("		ISNULL(#tmpRetenidoC.TotalRetenidoC,0),")
            loComandoSeleccionar.AppendLine("		#tmpExentoV.TotalExentoV,")
            loComandoSeleccionar.AppendLine("		SUM(#tmpBaseVFactura.BaseVFactura) + #tmpBaseVNotaCre.BaseVNotaCre AS TotalBaseV,")
            loComandoSeleccionar.AppendLine("		#tmpBaseVFactura.ImpuestoVFactura + #tmpBaseVNotaCre.ImpuestoVNotaCre AS TotalImpuestoV,")
            loComandoSeleccionar.AppendLine("		#tmpRetenidoV.TotalRetenidoV,")
            loComandoSeleccionar.AppendLine("       ISNULL((SELECT TotalVentaNoGravada FROM #tmpVentaNoGravada),0) AS TVNoGravada,")
            loComandoSeleccionar.AppendLine("       CAST(" & lcParametro1Desde & " AS DECIMAL(28,2))  AS ExcedenteC,")
            loComandoSeleccionar.AppendLine("       CAST(" & lcParametro2Desde & " AS DECIMAL(28,2))  AS Retenciones,")
            loComandoSeleccionar.AppendLine("       @sp_FecIni AS Desde, @sp_FecFin AS Hasta")
            loComandoSeleccionar.AppendLine("FROM #tmpExentoC,")
            loComandoSeleccionar.AppendLine("	#tmpBaseC,")
            loComandoSeleccionar.AppendLine(" 	#tmpBaseVFactura,")
            loComandoSeleccionar.AppendLine("	#tmpBaseVNotaCre,")
            loComandoSeleccionar.AppendLine(" 	#tmpExentoV,")
            loComandoSeleccionar.AppendLine("	#tmpRetenidoC,")
            loComandoSeleccionar.AppendLine("	#tmpRetenidoV")
            loComandoSeleccionar.AppendLine("GROUP BY #tmpExentoC.TotalExentoC, #tmpBaseC.TotalBaseC, #tmpBaseC.TotalImpuestoC, #tmpRetenidoC.TotalRetenidoC,#tmpExentoV.TotalExentoV, ")
            loComandoSeleccionar.AppendLine("		#tmpBaseVFactura.ImpuestoVFactura,#tmpBaseVNotaCre.BaseVNotaCre, #tmpBaseVNotaCre.ImpuestoVNotaCre, #tmpRetenidoV.TotalRetenidoV")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpBaseC")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpExentoC")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpBaseVFactura")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpBaseVNotaCre")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpExentoV")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpRetenidoC")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpRetenidoV")
            loComandoSeleccionar.AppendLine("DROP TABLE #tmpVentaNoGravada")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")
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


            '--------------------------------------------------'
            ' Carga la imagen del logo en cusReportes            '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rResumen_Iva", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTRF_rResumen_Iva.ReportSource = loObjetoReporte

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

    'Private Sub mGenerarArchivoExcel(ByVal lcNombreArchivo As String, ByVal loDatos As DataTable, ByVal lcParametrosReporte As String)

    '    Dim lnDecimalesMonto As Integer = goOpciones.pnDecimalesParaMonto
    '    Dim lnDecimalesCosto As Integer = goOpciones.pnDecimalesParaCosto
    '    Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
    '    Dim lnDecimalesPorcentaje As Integer = goOpciones.pnDecimalesParaPorcentaje

    '    Dim lcFormatoMontos As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesMonto)
    '    Dim lcFormatoCosto As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesCosto)

    '    Dim lcFormatoCantidad As String
    '    If (lnDecimalesCantidad > 0) Then
    '        lcFormatoCantidad = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesCantidad)
    '    Else
    '        lcFormatoCantidad = "###,###,###,###,##0"
    '    End If

    '    Dim lcFormatoPorcentaje As String = "###,###,###,###,##0." & Strings.left("0000000000", lnDecimalesPorcentaje)

    '    '******************************************************************'
    '    ' Declaración de objetos de excel: IMPORTANTE liberar recursos al	'
    '    ' final usando el GARBAGE COLLECTOR y ReleaseComObject.			'
    '    '******************************************************************'
    '    Dim loExcel As Excel.Application = Nothing
    '    Dim laLibros As Excel.Workbooks = Nothing
    '    Dim loLibro As Excel.Workbook = Nothing
    '    Dim loHoja As Excel.Worksheet = Nothing
    '    Dim loCeldas As Excel.Range = Nothing
    '    Dim loRango As Excel.Range = Nothing

    '    Dim loFilas As Excel.Range = Nothing
    '    Dim loColumnas As Excel.Range = Nothing
    '    Dim loFormas As Excel.Shapes = Nothing
    '    Dim loImagen As Excel.Shape = Nothing
    '    Dim loFuente As Excel.Font = Nothing


    '    Try

    '        ' Se inicializa el objeto de aplicacion excel
    '        loExcel = New Excel.Application()
    '        loExcel.Visible = False
    '        loExcel.DisplayAlerts = False

    '        ' Crea un nuevo libro de excel y activa la primera hoja
    '        laLibros = loExcel.Workbooks
    '        'loLibro = laLibros.Add()

    '        'Dim lcPlantilla As String = HttpContext.Current.Server.MapPath("~/Administrativo/Complementos/plantilla.xls")
    '        'System.IO.File.Copy(lcPlantilla, lcNombreArchivo)
    '        loLibro = laLibros.Open(lcNombreArchivo)

    '        loHoja = loLibro.Worksheets(1)
    '        loHoja.Activate()

    '        ' Formato por defecto de todas las celdas			
    '        loCeldas = loHoja.Range("A1:IV65536")
    '        'loCeldas = loHoja.Cells
    '        loCeldas.Clear()
    '        loFuente = loCeldas.Font
    '        loFuente.Size = 9
    '        loFuente.Name = "Tahoma"


    '        '******************************************************************'
    '        ' Encabezado de la hoja											'
    '        '******************************************************************'
    '        'Dim lcLogo As String = goEmpresa.pcUrlLogo 
    '        'lcLogo = HttpContext.Current.Server.MapPath(lcLogo)
    '        'loFormas = loHoja.Shapes

    '        'loFormas.AddPicture(lcLogo,  Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue, 1, 1, 60, 60)

    '        loRango = loHoja.Range("A1")
    '        loRango.Value = cusAplicacion.goEmpresa.pcNombre

    '        loRango = loHoja.Range("A2")
    '        loRango.Value = cusAplicacion.goEmpresa.pcRifEmpresa

    '        loRango = loHoja.Range("B5:T5")
    '        loRango.Select()
    '        loRango.MergeCells = True
    '        loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
    '        loRango.Value = "LIBRO DE COMPRAS"
    '        loFuente = loRango.Font
    '        loFuente.Size = 14
    '        loFuente.Bold = True

    '        'Sub título del reporte
    '        Dim ldFechaReporte As Date
    '        loRango = loHoja.Range("B6:T6")
    '        loRango.Select()
    '        loRango.MergeCells = True
    '        loRango.Value = "Mes de " & ldFechaReporte.ToString("MMMM 'de' yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("es-VE"))
    '        loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

    '        ' Fecha y hora de creacion
    '        Dim ldFecha As DateTime = Date.Now()
    '        loRango = loHoja.Range("T1")
    '        loRango.NumberFormat = "@"
    '        loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
    '        loRango.Value = ldFecha.ToString("dd/MM/yyyy")

    '        loRango = loHoja.Range("T2")
    '        loRango.NumberFormat = "@" 'La celda almacena un string
    '        loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
    '        loRango.Value = ldFecha.ToString("hh:mm:ss tt")

    '        ' Parametros del reporte
    '        'loRango = loHoja.Range("B7:O7")
    '        'loRango.Select()
    '        'loRango.MergeCells = True
    '        'loRango.Value = lcParametrosReporte
    '        'loRango.WrapText = True
    '        'loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify


    '        Dim lnFilaActual As Integer = 8

    '        '******************************************************************'
    '        ' Datos del Reporte												'
    '        '******************************************************************'
    '        loRango = loHoja.Range("B" & lnFilaActual)
    '        loRango.Value = "Oper." & vbLf & "Nro."

    '        loRango = loHoja.Range("C" & lnFilaActual)
    '        loRango.Value = "Fecha" & vbLf & "Contab."

    '        loRango = loHoja.Range("D" & lnFilaActual)
    '        loRango.Value = "Fecha de" & vbLf & "la Factura"

    '        loRango = loHoja.Range("E" & lnFilaActual)
    '        loRango.Value = "RIF"

    '        loRango = loHoja.Range("F" & lnFilaActual)
    '        loRango.Value = "Nombre o Razón Social"

    '        loRango = loHoja.Range("G" & lnFilaActual)
    '        loRango.Value = "Número" & vbLf & "Comprobante"

    '        loRango = loHoja.Range("H" & lnFilaActual)
    '        loRango.Value = "Núm. de" & vbLf & "Expediente" & vbLf & "Importación"

    '        loRango = loHoja.Range("I" & lnFilaActual)
    '        loRango.Value = "Número de" & vbLf & "Factura"

    '        loRango = loHoja.Range("J" & lnFilaActual)
    '        loRango.Value = "Número de" & vbLf & "Control de" & vbLf & "Factura"

    '        loRango = loHoja.Range("K" & lnFilaActual)
    '        loRango.Value = "Número de" & vbLf & "Nota de" & vbLf & "Débito"

    '        loRango = loHoja.Range("L" & lnFilaActual)
    '        loRango.Value = "Número de" & vbLf & "Nota de" & vbLf & "Crédito"

    '        loRango = loHoja.Range("M" & lnFilaActual)
    '        loRango.Value = "Tipo de" & vbLf & "Transac."

    '        loRango = loHoja.Range("N" & lnFilaActual)
    '        loRango.Value = "Número de" & vbLf & "Factura" & vbLf & "Afectada"

    '        loRango = loHoja.Range("O" & lnFilaActual)
    '        loRango.Value = "Total" & vbLf & "Compras" & vbLf & "Incl. IVA"

    '        loRango = loHoja.Range("P" & lnFilaActual)
    '        loRango.Value = "Compras" & vbLf & "sin drcho." & vbLf & "a IVA"

    '        loRango = loHoja.Range("Q" & lnFilaActual)
    '        loRango.Value = "Base" & vbLf & "Imponible"

    '        loRango = loHoja.Range("R" & lnFilaActual)
    '        loRango.Value = "%" & vbLf & "Alic."

    '        loRango = loHoja.Range("S" & lnFilaActual)
    '        loRango.Value = "Impuesto" & vbLf & "IVA"

    '        loRango = loHoja.Range("T" & lnFilaActual)
    '        loRango.Value = "IVA Retenido" & vbLf & "(por el vendedor)"

    '        loRango = loHoja.Range("B" & lnFilaActual & ":T" & lnFilaActual)
    '        loFuente = loRango.Font
    '        loFuente.Bold = True
    '        'loFuente.Color = Rgb(255, 255, 255)
    '        loRango.Interior.Color = Rgb(200, 200, 200)

    '        loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
    '        loRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
    '        loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

    '        '****************************************************************************************
    '        ' Facturas del Periodo actual
    '        '****************************************************************************************

    '        Dim lnFilaInicio As Integer = lnFilaActual
    '        Dim laRenglones() As DataRow = loDatos.Select("Periodo_Anterior=0")
    '        For Each loRenglon As DataRow In laRenglones
    '            'Dim loRenglon As DataRow = loDatos.Rows(lnRenglon)

    '            lnFilaActual += 1

    '            'Oper. Nro."
    '            loRango = loHoja.Range("B" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CInt(loRenglon("Operacion"))
    '            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

    '            'Fecha Contab.
    '            loRango = loHoja.Range("C" & lnFilaActual)
    '            loRango.NumberFormat = "dd-mm-yyyy;@"
    '            loRango.Value = CDate(loRenglon("Fec_Con"))
    '            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

    '            'Fecha de la Factura
    '            loRango = loHoja.Range("D" & lnFilaActual)
    '            loRango.NumberFormat = "dd-mm-yyyy;@"
    '            loRango.Value = CDate(loRenglon("Fec_Ini"))
    '            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

    '            'RIF
    '            loRango = loHoja.Range("E" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Rif")).Trim()

    '            'Nombre o Razón Social 
    '            loRango = loHoja.Range("F" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Nom_Pro")).Trim()

    '            'Número Comprobante
    '            loRango = loHoja.Range("G" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            If Not IsDBNull(loRenglon("Com_Ret")) Then
    '                loRango.Value = CStr(loRenglon("Com_Ret")).Trim()
    '            End If

    '            'Núm. de Expediente Importación
    '            loRango = loHoja.Range("H" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Expediente_Importacion")).Trim()

    '            'Número de Factura
    '            loRango = loHoja.Range("I" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Factura")).Trim()

    '            'Número de Control de Factura
    '            loRango = loHoja.Range("J" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Control")).Trim()

    '            'Número de Nota de Débito
    '            loRango = loHoja.Range("K" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Nota_Debito")).Trim()

    '            'Número de Nota de Crédito
    '            loRango = loHoja.Range("L" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Nota_Credito")).Trim()

    '            'Tipo de Transac.
    '            loRango = loHoja.Range("M" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Transaccion")).Trim()

    '            'Número de Factura Afectada
    '            loRango = loHoja.Range("N" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Documento_Afectado")).Trim()

    '            'Total Compras Incl. IVA
    '            loRango = loHoja.Range("O" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = CDec(loRenglon("Mon_Net"))

    '            'Compras sin drcho. a IVA
    '            loRango = loHoja.Range("P" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = CDec(loRenglon("Mon_Exe"))

    '            'Base Imponible
    '            loRango = loHoja.Range("Q" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = CDec(loRenglon("Mon_Bas"))

    '            '% Alic.
    '            Dim lnPorcentajeImpuesto As Decimal = CDec(loRenglon("Por_Imp"))
    '            loRango = loHoja.Range("R" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = lnPorcentajeImpuesto

    '            'Impuesto IVA
    '            loRango = loHoja.Range("S" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = CDec(loRenglon("Mon_Imp"))

    '            'IVA Retenido (por el vendedor)
    '            loRango = loHoja.Range("T" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = CDec(loRenglon("Mon_Ret"))

    '            'Condicion
    '            '   loRango = loHoja.Range("U" & lnFilaActual)
    '            'loRango.NumberFormat = "@"
    '            '   If (CStr(loRenglon("Status")).ToLower().Trim() = "anulado") Then
    '            '       loRango.Value = "ANULADO"
    '            '   Else 
    '            '       loRango.Value = IIf(cbool(loRenglon("Prov_Nacional")), "INTERNA", "IMPORTACION")
    '            '   End If

    '            'Alicuota
    '            loRango = loHoja.Range("U" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            If (CStr(loRenglon("Status")).ToLower().Trim() = "anulado") Then
    '                loRango.Value = "ANULADO"
    '            Else
    '                Dim lcTipoAlicuota As String
    '                lcTipoAlicuota = IIf(CBool(loRenglon("Prov_Nacional")), "INTERNA", "IMPORTACION")
    '                If (lnPorcentajeImpuesto = 0D) Then
    '                    lcTipoAlicuota = lcTipoAlicuota & "-EXENTO"
    '                ElseIf lnPorcentajeImpuesto < 12D Then
    '                    lcTipoAlicuota = lcTipoAlicuota & "-REDUCIDA"
    '                ElseIf lnPorcentajeImpuesto = 12D Then
    '                    lcTipoAlicuota = lcTipoAlicuota & "-GENERAL"
    '                Else 'If lnPorcentajeImpuesto > 12D 
    '                    lcTipoAlicuota = lcTipoAlicuota & "-ADICIONAL"
    '                End If
    '                loRango.Value = lcTipoAlicuota
    '            End If

    '        Next loRenglon

    '        Dim lnTotal As Integer = laRenglones.Length
    '        loRango = loHoja.Range("B" & (lnFilaInicio) & ":T" & (lnFilaInicio))
    '        loRango.Select()
    '        loExcel.Selection.AutoFilter()

    '        loRango = loHoja.Range("B" & (lnFilaInicio + 1) & ":T" & (lnFilaInicio + lnTotal))
    '        loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

    '        Dim lnDesde As Integer = lnFilaInicio
    '        Dim lnHasta As Integer = lnFilaInicio + lnTotal

    '        lnFilaInicio += lnTotal + 2
    '        loRango = loHoja.Range("B" & (lnFilaInicio) & ":C" & (lnFilaInicio))
    '        loRango.MergeCells = True
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Total Registros: " & lnTotal.ToString()
    '        loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

    '        loRango = loHoja.Range("N" & (lnFilaInicio))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Total General: "
    '        loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

    '        loRango = loHoja.Range("O" & (lnFilaInicio))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""<>ANULADO"", O" & lnDesde & ":O" & lnHasta & ")"

    '        loRango = loHoja.Range("P" & (lnFilaInicio))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""<>ANULADO"", P" & lnDesde & ":P" & lnHasta & ")"

    '        loRango = loHoja.Range("Q" & (lnFilaInicio))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""<>ANULADO"", Q" & lnDesde & ":Q" & lnHasta & ")"

    '        loRango = loHoja.Range("S" & (lnFilaInicio))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""<>ANULADO"", S" & lnDesde & ":S" & lnHasta & ")"

    '        loRango = loHoja.Range("T" & (lnFilaInicio))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""<>ANULADO"", T" & lnDesde & ":T" & lnHasta & ")"

    '        loRango = loHoja.Range("B" & (lnFilaInicio) & ":T" & (lnFilaInicio))
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        '****************************************************************************************
    '        ' Bloque de totales
    '        '****************************************************************************************
    '        lnFilaActual = lnFilaActual + 4
    '        loRango = loHoja.Range("K" & (lnFilaActual))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Base Imponible"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        loRango = loHoja.Range("L" & (lnFilaActual))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Credito Fiscal"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        loRango = loHoja.Range("M" & (lnFilaActual))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "IVA Retenido"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        lnFilaActual = lnFilaActual + 1
    '        loRango = loHoja.Range("G" & (lnFilaActual))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Total Compras Exentas y/o sin derecho a Crédito Fiscal"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        loRango = loHoja.Range("K" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""<>ANULADO"", P" & lnDesde & ":P" & lnHasta & ")"

    '        loRango = loHoja.Range("L" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = 0 '"=SUMIF(U" & lnDesde & ":U" & lnHasta	& ", ""<>ANULADO"", P" & lnDesde & ":P" & lnHasta	& ")"

    '        loRango = loHoja.Range("M" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = 0 '"=SUMIF(U" & lnDesde & ":U" & lnHasta	& ", ""<>ANULADO"", P" & lnDesde & ":P" & lnHasta	& ")"


    '        lnFilaActual = lnFilaActual + 1
    '        loRango = loHoja.Range("G" & (lnFilaActual))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Total Compras Importación Afectas solo Alícuota General"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        loRango = loHoja.Range("K" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=IMPORTACION-GENERAL"", Q" & lnDesde & ":Q" & lnHasta & ")"

    '        loRango = loHoja.Range("L" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=IMPORTACION-GENERAL"", S" & lnDesde & ":S" & lnHasta & ")"

    '        loRango = loHoja.Range("M" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=IMPORTACION-GENERAL"", T" & lnDesde & ":T" & lnHasta & ")"

    '        lnFilaActual = lnFilaActual + 1
    '        loRango = loHoja.Range("G" & (lnFilaActual))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Total Compras Importación Afectas en Alícuota General + Adicional"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        loRango = loHoja.Range("K" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=IMPORTACION-ADICIONAL"", Q" & lnDesde & ":Q" & lnHasta & ")"

    '        loRango = loHoja.Range("L" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=IMPORTACION-ADICIONAL"", S" & lnDesde & ":S" & lnHasta & ")"

    '        loRango = loHoja.Range("M" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=IMPORTACION-ADICIONAL"", T" & lnDesde & ":T" & lnHasta & ")"

    '        lnFilaActual = lnFilaActual + 1
    '        loRango = loHoja.Range("G" & (lnFilaActual))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Total Compras Importación Afectas en Alícuota Reducida"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        loRango = loHoja.Range("K" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=IMPORTACION-REDUCIDA"", Q" & lnDesde & ":Q" & lnHasta & ")"

    '        loRango = loHoja.Range("L" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=IMPORTACION-REDUCIDA"", S" & lnDesde & ":S" & lnHasta & ")"

    '        loRango = loHoja.Range("M" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=IMPORTACION-REDUCIDA"", T" & lnDesde & ":T" & lnHasta & ")"


    '        lnFilaActual = lnFilaActual + 1
    '        loRango = loHoja.Range("G" & (lnFilaActual))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Total Compras Internas Afectas solo Alícuota General"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        loRango = loHoja.Range("K" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=INTERNA-GENERAL"", Q" & lnDesde & ":Q" & lnHasta & ")"

    '        loRango = loHoja.Range("L" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=INTERNA-GENERAL"", S" & lnDesde & ":S" & lnHasta & ")"

    '        loRango = loHoja.Range("M" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=INTERNA-GENERAL"", T" & lnDesde & ":T" & lnHasta & ")"

    '        lnFilaActual = lnFilaActual + 1
    '        loRango = loHoja.Range("G" & (lnFilaActual))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Total Compras Internas Afectas en Alícuota General + Adicional"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        loRango = loHoja.Range("K" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=INTERNA-ADICIONAL"", Q" & lnDesde & ":Q" & lnHasta & ")"

    '        loRango = loHoja.Range("L" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=INTERNA-ADICIONAL"", S" & lnDesde & ":S" & lnHasta & ")"

    '        loRango = loHoja.Range("M" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=INTERNA-ADICIONAL"", T" & lnDesde & ":T" & lnHasta & ")"

    '        lnFilaActual = lnFilaActual + 1
    '        loRango = loHoja.Range("G" & (lnFilaActual))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Total Compras Internas Afectas en Alícuota Reducida"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        loRango = loHoja.Range("K" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=INTERNA-REDUCIDA"", Q" & lnDesde & ":Q" & lnHasta & ")"

    '        loRango = loHoja.Range("L" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=INTERNA-REDUCIDA"", S" & lnDesde & ":S" & lnHasta & ")"

    '        loRango = loHoja.Range("M" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""=INTERNA-REDUCIDA"", T" & lnDesde & ":T" & lnHasta & ")"

    '        lnFilaActual = lnFilaActual + 1
    '        lnDesde = lnFilaActual - 7
    '        lnHasta = lnFilaActual - 1

    '        loRango = loHoja.Range("K" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUM(K" & lnDesde & ":K" & lnHasta & ")"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        loRango = loHoja.Range("L" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUM(L" & lnDesde & ":L" & lnHasta & ")"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        loRango = loHoja.Range("M" & (lnFilaActual))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUM(M" & lnDesde & ":M" & lnHasta & ")"
    '        loFuente = loRango.Font
    '        loFuente.Bold = True


    '        '****************************************************************************************
    '        ' Facturas del Periodo anterior
    '        '****************************************************************************************
    '        lnFilaActual = lnFilaActual + 3

    '        loRango = loHoja.Range("B" & lnFilaActual)
    '        loFuente = loRango.Font
    '        loFuente.Bold = True
    '        loFuente.Size = 14
    '        loRango.Value = "AJUSTES"

    '        lnFilaActual = lnFilaActual + 1

    '        loRango = loHoja.Range("B" & lnFilaActual)
    '        loRango.Value = "Oper." & vbLf & "Nro."

    '        loRango = loHoja.Range("C" & lnFilaActual)
    '        loRango.Value = "Fecha" & vbLf & "Contab."

    '        loRango = loHoja.Range("D" & lnFilaActual)
    '        loRango.Value = "Fecha de" & vbLf & "la Factura"

    '        loRango = loHoja.Range("E" & lnFilaActual)
    '        loRango.Value = "RIF"

    '        loRango = loHoja.Range("F" & lnFilaActual)
    '        loRango.Value = "Nombre o Razón Social"

    '        loRango = loHoja.Range("G" & lnFilaActual)
    '        loRango.Value = "Número" & vbLf & "Comprobante"

    '        loRango = loHoja.Range("H" & lnFilaActual)
    '        loRango.Value = "Núm. de" & vbLf & "Expediente" & vbLf & "Importación"

    '        loRango = loHoja.Range("I" & lnFilaActual)
    '        loRango.Value = "Número de" & vbLf & "Factura"

    '        loRango = loHoja.Range("J" & lnFilaActual)
    '        loRango.Value = "Número de" & vbLf & "Control de" & vbLf & "Factura"

    '        loRango = loHoja.Range("K" & lnFilaActual)
    '        loRango.Value = "Número de" & vbLf & "Nota de" & vbLf & "Débito"

    '        loRango = loHoja.Range("L" & lnFilaActual)
    '        loRango.Value = "Número de" & vbLf & "Nota de" & vbLf & "Crédito"

    '        loRango = loHoja.Range("M" & lnFilaActual)
    '        loRango.Value = "Tipo de" & vbLf & "Transac."

    '        loRango = loHoja.Range("N" & lnFilaActual)
    '        loRango.Value = "Número de" & vbLf & "Factura" & vbLf & "Afectada"

    '        loRango = loHoja.Range("O" & lnFilaActual)
    '        loRango.Value = "Total" & vbLf & "Compras" & vbLf & "Incl. IVA"

    '        loRango = loHoja.Range("P" & lnFilaActual)
    '        loRango.Value = "Compras" & vbLf & "sin drcho." & vbLf & "a IVA"

    '        loRango = loHoja.Range("Q" & lnFilaActual)
    '        loRango.Value = "Base" & vbLf & "Imponible"

    '        loRango = loHoja.Range("R" & lnFilaActual)
    '        loRango.Value = "%" & vbLf & "Alic."

    '        loRango = loHoja.Range("S" & lnFilaActual)
    '        loRango.Value = "Impuesto" & vbLf & "IVA"

    '        loRango = loHoja.Range("T" & lnFilaActual)
    '        loRango.Value = "IVA Retenido" & vbLf & "(por el vendedor)"

    '        loRango = loHoja.Range("B" & lnFilaActual & ":T" & lnFilaActual)
    '        loFuente = loRango.Font
    '        loFuente.Bold = True
    '        'loFuente.Color = Rgb(255, 255, 255)
    '        loRango.Interior.Color = Rgb(200, 200, 200)

    '        loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
    '        loRango.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
    '        loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)


    '        lnFilaInicio = lnFilaActual
    '        laRenglones = loDatos.Select("Periodo_Anterior=1")
    '        For Each loRenglon As DataRow In laRenglones

    '            lnFilaActual += 1

    '            'Oper. Nro."
    '            loRango = loHoja.Range("B" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CInt(loRenglon("Operacion"))
    '            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

    '            'Fecha Contab.
    '            loRango = loHoja.Range("C" & lnFilaActual)
    '            loRango.NumberFormat = "dd-mm-yyyy;@"
    '            loRango.Value = CDate(loRenglon("Fec_Con"))
    '            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

    '            'Fecha de la Factura
    '            loRango = loHoja.Range("D" & lnFilaActual)
    '            loRango.NumberFormat = "dd-mm-yyyy;@"
    '            'loRango.Value = CDate(loRenglon("Fec_Ini"))
    '            loRango.Value = CDate(loRenglon("Fec_Doc"))
    '            loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter

    '            'RIF
    '            loRango = loHoja.Range("E" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Rif")).Trim()

    '            'Nombre o Razón Social 
    '            loRango = loHoja.Range("F" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Nom_Pro")).Trim()

    '            'Número Comprobante
    '            loRango = loHoja.Range("G" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            If Not IsDBNull(loRenglon("Com_Ret")) Then
    '                loRango.Value = CStr(loRenglon("Com_Ret")).Trim()
    '            End If

    '            'Núm. de Expediente Importación
    '            loRango = loHoja.Range("H" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Expediente_Importacion")).Trim()

    '            'Número de Factura
    '            loRango = loHoja.Range("I" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Factura")).Trim()

    '            'Número de Control de Factura
    '            loRango = loHoja.Range("J" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Control")).Trim()

    '            'Número de Nota de Débito
    '            loRango = loHoja.Range("K" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Nota_Debito")).Trim()

    '            'Número de Nota de Crédito
    '            loRango = loHoja.Range("L" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Nota_Credito")).Trim()

    '            'Tipo de Transac.
    '            loRango = loHoja.Range("M" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Transaccion")).Trim()

    '            'Número de Factura Afectada
    '            loRango = loHoja.Range("N" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            loRango.Value = CStr(loRenglon("Documento_Afectado")).Trim()

    '            'Total Compras Incl. IVA
    '            loRango = loHoja.Range("O" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = CDec(loRenglon("Mon_Net"))

    '            'Compras sin drcho. a IVA
    '            loRango = loHoja.Range("P" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = CDec(loRenglon("Mon_Exe"))

    '            'Base Imponible
    '            loRango = loHoja.Range("Q" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = CDec(loRenglon("Mon_Bas"))

    '            '% Alic.
    '            Dim lnPorcentajeImpuesto As Decimal = CDec(loRenglon("Por_Imp"))
    '            loRango = loHoja.Range("R" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = lnPorcentajeImpuesto

    '            'Impuesto IVA
    '            loRango = loHoja.Range("S" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = CDec(loRenglon("Mon_Imp"))

    '            'IVA Retenido (por el vendedor)
    '            loRango = loHoja.Range("T" & lnFilaActual)
    '            loRango.NumberFormat = lcFormatoMontos
    '            loRango.Value = CDec(loRenglon("Mon_Ret"))

    '            'Alicuota
    '            loRango = loHoja.Range("U" & lnFilaActual)
    '            loRango.NumberFormat = "@"
    '            If (CStr(loRenglon("Status")).ToLower().Trim() = "anulado") Then
    '                loRango.Value = "ANULADO"
    '            Else
    '                Dim lcTipoAlicuota As String
    '                lcTipoAlicuota = IIf(CBool(loRenglon("Prov_Nacional")), "INTERNA", "IMPORTACION")
    '                If (lnPorcentajeImpuesto = 0D) Then
    '                    lcTipoAlicuota = lcTipoAlicuota & "-EXENTO"
    '                ElseIf lnPorcentajeImpuesto < 12D Then
    '                    lcTipoAlicuota = lcTipoAlicuota & "-REDUCIDA"
    '                ElseIf lnPorcentajeImpuesto = 12D Then
    '                    lcTipoAlicuota = lcTipoAlicuota & "-GENERAL"
    '                Else 'If lnPorcentajeImpuesto > 12D 
    '                    lcTipoAlicuota = lcTipoAlicuota & "-ADICIONAL"
    '                End If
    '                loRango.Value = lcTipoAlicuota
    '            End If

    '        Next loRenglon


    '        lnTotal = laRenglones.Length
    '        'loRango = loHoja.Range("B" & (lnFilaInicio) & ":T" & (lnFilaInicio))
    '        'loRango.Select() 
    '        'loExcel.Selection.AutoFilter()

    '        loRango = loHoja.Range("B" & (lnFilaInicio + 1) & ":T" & (lnFilaInicio + lnTotal))
    '        loRango.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium)

    '        lnDesde = lnFilaInicio
    '        lnHasta = lnFilaInicio + lnTotal

    '        lnFilaInicio += lnTotal + 2
    '        loRango = loHoja.Range("B" & (lnFilaInicio) & ":C" & (lnFilaInicio))
    '        loRango.MergeCells = True
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Total Registros: " & lnTotal.ToString()
    '        loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft

    '        loRango = loHoja.Range("N" & (lnFilaInicio))
    '        loRango.NumberFormat = "@"
    '        loRango.Value = "Total General: "
    '        loRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight

    '        loRango = loHoja.Range("O" & (lnFilaInicio))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""<>ANULADO"", O" & lnDesde & ":O" & lnHasta & ")"

    '        loRango = loHoja.Range("P" & (lnFilaInicio))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""<>ANULADO"", P" & lnDesde & ":P" & lnHasta & ")"

    '        loRango = loHoja.Range("Q" & (lnFilaInicio))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""<>ANULADO"", Q" & lnDesde & ":Q" & lnHasta & ")"

    '        loRango = loHoja.Range("S" & (lnFilaInicio))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""<>ANULADO"", S" & lnDesde & ":S" & lnHasta & ")"

    '        loRango = loHoja.Range("T" & (lnFilaInicio))
    '        loRango.NumberFormat = lcFormatoMontos
    '        loRango.Formula = "=SUMIF(U" & lnDesde & ":U" & lnHasta & ", ""<>ANULADO"", T" & lnDesde & ":T" & lnHasta & ")"

    '        loRango = loHoja.Range("B" & (lnFilaInicio) & ":T" & (lnFilaInicio))
    '        loFuente = loRango.Font
    '        loFuente.Bold = True

    '        '****************************************************************************************
    '        ' Ajustes finales de formato (tamaño de celdas, etc...)
    '        '****************************************************************************************
    '        loFilas = loCeldas.Rows
    '        loFilas.AutoFit()

    '        loColumnas = loCeldas.Columns
    '        loColumnas.AutoFit()

    '        loRango = loHoja.Range("A1:A" & lnFilaInicio)
    '        loRango.ColumnWidth = 2

    '        loRango = loHoja.Range("B1:B" & lnFilaInicio)
    '        loRango.ColumnWidth = 6

    '        loRango = loHoja.Range("C1:C" & lnFilaInicio)
    '        loRango.ColumnWidth = 11

    '        loRango = loHoja.Range("D1:D" & lnFilaInicio)
    '        loRango.ColumnWidth = 11

    '        loRango = loHoja.Range("E1:E" & lnFilaInicio)
    '        loRango.ColumnWidth = 14

    '        loRango = loHoja.Range("F1:F" & lnFilaInicio)
    '        loRango.ColumnWidth = 35

    '        loRango = loHoja.Range("G1:G" & lnFilaInicio)
    '        loRango.ColumnWidth = 18

    '        loRango = loHoja.Range("H1:H" & lnFilaInicio)
    '        loRango.ColumnWidth = 13

    '        loRango = loHoja.Range("I1:I" & lnFilaInicio)
    '        loRango.ColumnWidth = 13

    '        loRango = loHoja.Range("J1:J" & lnFilaInicio)
    '        loRango.ColumnWidth = 16

    '        loRango = loHoja.Range("K1:K" & lnFilaInicio)
    '        loRango.ColumnWidth = 13

    '        loRango = loHoja.Range("L1:L" & lnFilaInicio)
    '        loRango.ColumnWidth = 13

    '        loRango = loHoja.Range("M1:M" & lnFilaInicio)
    '        loRango.ColumnWidth = 10

    '        loRango = loHoja.Range("N1:N" & lnFilaInicio)
    '        loRango.ColumnWidth = 13

    '        loRango = loHoja.Range("O1:Q" & lnFilaInicio)
    '        loRango.ColumnWidth = 14

    '        loRango = loHoja.Range("R1:R" & lnFilaInicio)
    '        loRango.ColumnWidth = 11

    '        loRango = loHoja.Range("S1:U" & lnFilaInicio)
    '        loRango.ColumnWidth = 14

    '        ' Seleccionamos la primera celda del libro
    '        loRango = loHoja.Range("A1")
    '        loRango.Select()

    '        'Guardamos los cambios del libro activo
    '        loLibro.SaveAs(lcNombreArchivo)

    '        '******************************************************************'
    '        ' IMPORTANTE: Forma correcta de liberar recursos!!!				'
    '        '******************************************************************'
    '        ' Cerramos y liberamos recursos

    '    Catch loExcepcion As Exception

    '        Throw New Exception("No fue posible exportar los datos a excel. " & loExcepcion.Message, loExcepcion)

    '    Finally

    '        If (loFuente IsNot Nothing) Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(loFuente)
    '            loFuente = Nothing
    '        End If

    '        If (loFormas IsNot Nothing) Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(loFormas)
    '            loFormas = Nothing
    '        End If

    '        If (loRango IsNot Nothing) Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(loRango)
    '            loRango = Nothing
    '        End If

    '        If (loFilas IsNot Nothing) Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(loFilas)
    '            loFilas = Nothing
    '        End If

    '        If (loColumnas IsNot Nothing) Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(loColumnas)
    '            loColumnas = Nothing
    '        End If

    '        If (loCeldas IsNot Nothing) Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(loCeldas)
    '            loCeldas = Nothing
    '        End If

    '        If (loHoja IsNot Nothing) Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(loHoja)
    '            loHoja = Nothing
    '        End If

    '        If (loLibro IsNot Nothing) Then
    '            loLibro.Close(True)
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(loLibro)
    '            loLibro = Nothing
    '        End If

    '        If (laLibros IsNot Nothing) Then
    '            System.Runtime.InteropServices.Marshal.ReleaseComObject(laLibros)
    '            laLibros = Nothing
    '        End If

    '        loExcel.Quit()

    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(loExcel)
    '        loExcel = Nothing

    '        GC.Collect()
    '        GC.WaitForPendingFinalizers()

    '    End Try

    'End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 06/10/14: Codigo inicial.					                                        '
'-------------------------------------------------------------------------------------------'
' RJG: 17/03/15: Se eliminó el campo "Fecha de contabilización" y se agregó el número de    '
'                factura del proveedor.					                                    '
'-------------------------------------------------------------------------------------------'
' RJG: 29/04/15: Se ampliaron los campos de Factura y Control. Se ajustó el ordenamiento del'
'                grupo de ajustes. Se agregó el "impuesto excluido" de la GACETA 6152.		'
'-------------------------------------------------------------------------------------------'
