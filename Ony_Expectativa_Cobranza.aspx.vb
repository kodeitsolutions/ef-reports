'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
Imports System
Imports System.Collections.Specialized
Imports System.Net

'-------------------------------------------------------------------------------------------'
' Inicio de clase "Ony_Expectativa_Cobranza"
'-------------------------------------------------------------------------------------------'
Partial Class Ony_Expectativa_Cobranza
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument
    Dim loAppExcel As Excel.Application


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro11Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro12Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(12), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro12Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(12), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro13Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(13), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro13Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(13), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT Cuentas_Cobrar.Documento, ")
            loComandoSeleccionar.AppendLine(" CAST(Cuentas_Cobrar.Seg_Adm as XML) AS Fechas, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Tip, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Fec_Fin, ")
            'loComandoSeleccionar.AppendLine("           DATEDIFF(dd,Cuentas_Cobrar.Fec_Fin,GETDATE()) as Dias, ")
            loComandoSeleccionar.AppendLine(" CASE  	WHEN DATEDIFF(dd,Cuentas_Cobrar.Fec_Fin,GETDATE())>0 then DATEDIFF(dd,Cuentas_Cobrar.Fec_Fin,GETDATE()) Else 0 END As Dias, ")
            

            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 			Clientes.Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Tra, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Control, ")
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Cobrar.Mon_Bru *(-1) Else Cuentas_Cobrar.Mon_Bru End) As Mon_Bru, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar.Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Cobrar.Mon_Net *(-1) Else Cuentas_Cobrar.Mon_Net End) As Mon_Net, ")
            loComandoSeleccionar.AppendLine("           (Case when Tip_Doc = 'Credito' then Cuentas_Cobrar.Mon_Sal *(-1) Else Cuentas_Cobrar.Mon_Sal End) As Mon_Sal,  ")
            loComandoSeleccionar.AppendLine(" 			Vendedores.Nom_Ven  ")
            loComandoSeleccionar.AppendLine(" INTO      #tmp ")
            loComandoSeleccionar.AppendLine(" FROM      Clientes, ")
            loComandoSeleccionar.AppendLine(" 			Cuentas_Cobrar, ")
            loComandoSeleccionar.AppendLine(" 			Vendedores, ")
            loComandoSeleccionar.AppendLine(" 			Transportes, ")
            loComandoSeleccionar.AppendLine(" 			Monedas ")
            loComandoSeleccionar.AppendLine(" WHERE Cuentas_Cobrar.Cod_Cli = Clientes.Cod_Cli ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Tra = Transportes.Cod_Tra ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Mon = Monedas.Cod_Mon ")
            loComandoSeleccionar.AppendLine(" 			AND Cuentas_Cobrar.Cod_Ven = Vendedores.Cod_Ven ")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Documento    Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Fec_Ini      Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tip      Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Cli      Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Ven      Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Zon      Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Status       IN ( " & lcParametro6Desde & ")")
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Tip    Between " & lcParametro7Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("           And Clientes.Cod_Cla      Between " & lcParametro8Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro8Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Tra      Between " & lcParametro9Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro9Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Mon      Between " & lcParametro10Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro10Hasta)
            loComandoSeleccionar.AppendLine("		    AND ((" & lcParametro11Desde & " = 'Si' AND Cuentas_Cobrar.Mon_Sal > 0)")
            loComandoSeleccionar.AppendLine("			OR (" & lcParametro11Desde & " <> 'Si' AND (Cuentas_Cobrar.Mon_Sal >= 0 or Cuentas_Cobrar.Mon_Sal < 0)))")
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Suc      Between " & lcParametro12Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro12Hasta)
            loComandoSeleccionar.AppendLine("           And Cuentas_Cobrar.Cod_Rev      Between " & lcParametro13Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro13Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   Cuentas_Cobrar.Cod_Tip," & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("           ")
            loComandoSeleccionar.AppendLine(" SELECT	Documento, ")
            loComandoSeleccionar.AppendLine("           D.C.value('@fecha', 'datetime')			    as Fecha_Estimada,       ")
            loComandoSeleccionar.AppendLine("           D.C.value('@accion', 'Varchar(100)')		as Accion       ")
            loComandoSeleccionar.AppendLine(" INTO      #tmp2          ")
            loComandoSeleccionar.AppendLine(" FROM      #tmp          ")
            loComandoSeleccionar.AppendLine("           CROSS APPLY Fechas.nodes('elementos/elemento') D(c)           ")
            loComandoSeleccionar.AppendLine(" ORDER BY  Documento ASC ,Fecha_Estimada Desc            ")
            loComandoSeleccionar.AppendLine("           ")
            loComandoSeleccionar.AppendLine(" SELECT    Documento,  ")
            loComandoSeleccionar.AppendLine("           Fecha_Estimada,  ")
            loComandoSeleccionar.AppendLine("           Accion          ")
            loComandoSeleccionar.AppendLine(" INTO      #tmp3          ")
            loComandoSeleccionar.AppendLine(" FROM      #tmp2          ")
            loComandoSeleccionar.AppendLine(" WHERE     Fecha_Estimada          ")
            loComandoSeleccionar.AppendLine(" IN        (SELECT MAX(Fecha_Estimada) FROM #tmp2 GROUP BY Documento)          ")
            loComandoSeleccionar.AppendLine("           ")
            loComandoSeleccionar.AppendLine(" SELECT    #tmp.Documento          AS Documento, ")
            loComandoSeleccionar.AppendLine(" 			#tmp.Cod_Tip            AS Cod_Tip, ")
            loComandoSeleccionar.AppendLine(" 			#tmp.Fec_Ini            AS Fec_Ini, ")
            loComandoSeleccionar.AppendLine(" 			#tmp.Fec_Fin            AS Fec_Fin, ")
            loComandoSeleccionar.AppendLine(" 			#tmp.Dias               AS Dias, ")
            loComandoSeleccionar.AppendLine(" 			#tmp.Cod_Cli            AS Cod_Cli, ")
            loComandoSeleccionar.AppendLine(" 			#tmp.Nom_Cli            AS Nom_Cli, ")
            loComandoSeleccionar.AppendLine(" 			#tmp.Cod_Ven            AS Cod_Ven, ")
            loComandoSeleccionar.AppendLine(" 			#tmp.Cod_Tra            AS Cod_Tra, ")
            loComandoSeleccionar.AppendLine(" 			#tmp.Cod_Mon            AS Cod_Mon, ")
            loComandoSeleccionar.AppendLine(" 			#tmp.Control            AS Control, ")
            loComandoSeleccionar.AppendLine("           #tmp.Mon_Bru            AS Mon_Bru, ")
            loComandoSeleccionar.AppendLine(" 			#tmp.Mon_Imp1           AS Mon_Imp1, ")
            loComandoSeleccionar.AppendLine("           #tmp.Mon_Net            AS Mon_Net, ")
            loComandoSeleccionar.AppendLine("           #tmp.Mon_Sal            AS Mon_Sal,  ")
            loComandoSeleccionar.AppendLine(" 		    #tmp.Nom_Ven            AS Nom_Ven, ")
            loComandoSeleccionar.AppendLine("           #tmp3.Fecha_Estimada    AS Fecha_Estimada, ")
            loComandoSeleccionar.AppendLine("           ISNULL(#tmp3.Accion,'') AS Accion ")
            loComandoSeleccionar.AppendLine(" FROM      #tmp          ")
            loComandoSeleccionar.AppendLine(" LEFT JOIN #tmp3 ON #tmp3.documento = #tmp.documento          ")



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

			'-------------------------------------------------------------------
            ' Selección de opcion por excel (Microsoft Excel - xls)
			'-------------------------------------------------------------------
            If (Me.Request.QueryString("salida").ToLower() = "xls") Then
                ' Genera el archivo a partir de la tabla de datos y termina la ejecución
                Me.mGenerarArchivoExcel(laDatosReporte)

            End If

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("Ony_Expectativa_Cobranza", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvOny_Expectativa_Cobranza.ReportSource = loObjetoReporte

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
    ''' Rutina que modifica y formatea el contenido del archivo excel a descargar.
    ''' </summary>
    ''' <param name="loFileName">Ruta del archivo a modificar.</param>
    ''' <param name="loDatosReporte">Datos de la consulta a la base de datos.</param>
    ''' <remarks></remarks>
    Private Sub mGenerarArchivoExcel(ByVal loDatos As DataSet)

    '***********************************************************************'
    ' Prepara los datos para enviarlos al servicio web de Excel.            '
    '***********************************************************************'
        Dim loSalida As New IO.MemoryStream()
        loDatos.WriteXml(loSalida, XmlWriteMode.WriteSchema)


    '***********************************************************************'
    ' Prepara los parámetros adicionales para enviarlos junto con los datos.'
    '***********************************************************************'
        Dim lnDecimalesMonto As Integer = goOpciones.pnDecimalesParaMonto
        Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
        Dim lnDecimalesPorcentaje As Integer = goOpciones.pnDecimalesParaPorcentaje


        Dim loParametros As New NameValueCollection()
        loParametros.Add("lcNombreEmpresa", cusAplicacion.goEmpresa.pcNombre)
        loParametros.Add("lcRifEmpresa", "") 'cusAplicacion.goEmpresa.pcRifEmpresa)
        loParametros.Add("lnDecimalesMonto", lnDecimalesMonto.ToString())
        loParametros.Add("lnDecimalesCantidad", lnDecimalesCantidad.ToString())
        loParametros.Add("lnDecimalesPorcentaje", lnDecimalesPorcentaje.ToString())

        Dim loClienteWeb As new WebClient()
        loClienteWeb.QueryString = loParametros
        loClienteWeb.Headers.Add("Cache-Control", "no-cache")
    '***********************************************************************'
    ' Envía los datos y parámetros, y espera la respuesta.                  '
    '***********************************************************************'
        Dim loRespuesta As Byte()  
        Try
            loRespuesta = loClienteWeb.UploadData("http://localhost:8010/Reportes/Ony_Expectativa_Cobranza_xlsx.aspx", loSalida.GetBuffer())
        Catch ex As Exception
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado" , _ 
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
                                                                 ex.Message, vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return
        End Try

    '***********************************************************************'
    ' Vemos si la respuesta es TextoPlano (error) o no (el archivo Excel    '
    ' generado). Si el tipo está vacio : error desconocido.                 '
    '***********************************************************************'
        Dim loTipoRespuesta As String = loClienteWeb.ResponseHeaders("Content-Type") 

        If String.IsNullOrEmpty(loTipoRespuesta) Then 
            'Error no especificado!
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado" , _ 
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: El servicio que genera la salida XSLX no responde.<br/>", _
                                                                 vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return

        ElseIf loTipoRespuesta.ToLower().StartsWith("text/plain") Then 

            Dim lcMensaje As String = UTF32Encoding.UTF8.GetString(loRespuesta)
            lcMensaje = Me.Server.HtmlEncode(lcMensaje)

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado" , _ 
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
                                                                 lcMensaje, vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return

        Else
            'Generación exitosa: la respuesta es el archivo en excel para descargar

            Me.Response.Clear()
            Me.Response.Buffer = True
            Me.Response.AppendHeader("content-disposition", "attachment; filename=Ony_Expectativa_Cobranza.xlsx")
            Me.Response.ContentType = "application/excel"
            Me.Response.BinaryWrite(loRespuesta)
            Me.Response.End()

        End If

    End Sub


    ''' <summary>
    ''' Rutina que modifica y formatea el contenido del archivo excel a descargar.
    ''' </summary>
    ''' <param name="loFileName">Ruta del archivo a modificar.</param>
    ''' <param name="loDatosReporte">Datos de la consulta a la base de datos.</param>
    ''' <remarks></remarks>
    Private Sub modificar_excel(ByVal loFileName As String, ByVal loDatosReporte As DataSet)

        'Try
        '    ' Se inicializa el objeto a la aplicacion excel
        '    loAppExcel = New Excel.Application()
        '    loAppExcel.Visible = False
        '    loAppExcel.DisplayAlerts = False

        '    ' Se carga el archivo excel a modificar
        '    Dim lcLibrosExcel As Excel.Workbooks = loAppExcel.Workbooks
        '    Dim lcLibroExcel As Excel.Workbook = lcLibrosExcel.Open(loFileName)

        '    ' Se activa la primera hoja del libro donde se almacenara toda la informacion
        '    Dim lcHojaExcel As Excel.Worksheet
        '    lcHojaExcel = lcLibroExcel.Worksheets(1)
        '    lcHojaExcel.Activate()

        '    ' Se selecciona toda la hoja para blanquera todo
        '    '   - El número total de columnas disponibles en Excel
        '    '       Viejo límite: 256 (2^8)         (Excel 2003 o inferor)
        '    '       Nuevo límite: 16.384 (2^14)     (Excel 2007)
        '    '   - El número total de filas disponibles en Excel
        '    '       Viejo límite: 65.536 (2^16)         (Excel 2003 o inferior)
        '    '       Nuevo límite: el 1.048.576 (2^20)   (Excel 2007)
        '    Dim lcRango As Excel.Range = lcHojaExcel.Range("A1:IV65536")
        '    lcRango.Select()
        '    lcRango.Clear()
        '    lcRango.Font.Size = 8
        '    lcRango.Font.Name = "Tahoma"

        '    ' Nombre de la empresa
        '    lcHojaExcel.Cells(1, 1).Value = cusAplicacion.goEmpresa.pcNombre.ToUpper
        '    ' Nombre del modulo
        '    lcHojaExcel.Cells(2, 1).Value = "Cuentas x Cobrar"
        '    ' Titulo del reporte
        '    lcRango = lcHojaExcel.Range("A3:P3")
        '    lcRango.Select()
        '    lcRango.Font.ColorIndex = 25
        '    lcRango.Interior.ColorIndex = 34
        '    lcRango.MergeCells = True
        '    lcRango.Value = Me.pcNombreReporte
        '    lcRango.Font.Size = 14
        '    lcRango.Font.Bold = True
        '    lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        '    lcRango.Rows.AutoFit()

        '    ' Parametros del reporte
        '    lcRango = lcHojaExcel.Range("A4:P4")
        '    lcRango.Select()
        '    lcRango.MergeCells = True
        '    lcRango.Value = cusAplicacion.goReportes.mObtenerParametros(cusAplicacion.goReportes.paNombresParametros, cusAplicacion.goReportes.paParametrosIniciales, cusAplicacion.goReportes.paParametrosFinales)
        '    lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignJustify
        '    lcRango.Rows.AutoFit()


        '    lcHojaExcel.Cells(6, 1).Value = "Tipo"
        '    lcHojaExcel.Cells(6, 2).Value = "Código"
        '    lcHojaExcel.Cells(6, 3).Value = "Nombre"
        '    lcHojaExcel.Cells(6, 4).Value = "Documento"
        '    lcHojaExcel.Cells(6, 5).Value = "Control"
        '    lcHojaExcel.Cells(6, 6).Value = "Emision"
        '    lcHojaExcel.Cells(6, 7).Value = "Vencimiento"
        '    lcHojaExcel.Cells(6, 8).Value = "Dias Vencido"
        '    lcHojaExcel.Cells(6, 9).Value = "Fecha Estimada Pago"
        '    lcHojaExcel.Cells(6, 10).Value = "Vendedor"
        '    lcHojaExcel.Cells(6, 11).Value = "Moneda"
        '    lcHojaExcel.Cells(6, 12).Value = "Monto Bruto"
        '    lcHojaExcel.Cells(6, 13).Value = "Monto Impuesto"
        '    lcHojaExcel.Cells(6, 14).Value = "Monto Neto"
        '    lcHojaExcel.Cells(6, 15).Value = "Saldo"
        '    lcHojaExcel.Cells(6, 16).Value = "Observación"
            
        '    ' Se le da formato a las celdas del membrete de la tabla

        '    lcRango = lcHojaExcel.Range("A6:P6")
        '    lcRango.Select()
        '    lcRango.Font.Bold = True
        '    lcRango.EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        '    lcRango.EntireRow.WrapText = True
        '    lcRango.Font.ColorIndex = 34
        '    lcRango.Interior.ColorIndex = 25

        '    ' Formato a las columnas de la tabla de informacion del reporte

        '    lcRango = lcHojaExcel.Range("A1:E1")
        '    lcRango.EntireColumn.NumberFormat = "@"

        '    lcRango = lcHojaExcel.Range("F1:G1")
        '    lcRango.EntireColumn.NumberFormat = "DD/MM/YYYY"

        '    lcRango = lcHojaExcel.Range("H1")
        '    lcRango.EntireColumn.NumberFormat = "#####0"

        '    lcRango = lcHojaExcel.Range("I1")
        '    lcRango.EntireColumn.NumberFormat = "DD/MM/YYYY"


        '    lcRango = lcHojaExcel.Range("J1:K1")
        '    lcRango.EntireColumn.NumberFormat = "@"

        '    lcRango = lcHojaExcel.Range("L1:O1")
        '    lcRango.EntireColumn.NumberFormat = "###,###,##0"

        '    lcRango = lcHojaExcel.Range("P1")
        '    lcRango.EntireColumn.NumberFormat = "@"

        '    ' Formato a las celdas de la fecha y hora de creacion
        '    ' Fecha y hora de creacion
        '    Dim lcFechaCreacion As DateTime = DateTime.Now()
        '    lcHojaExcel.Cells(1, 16).NumberFormat = "@"
        '    lcHojaExcel.Cells(1, 16).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        '    lcHojaExcel.Cells(1, 16).Value = lcFechaCreacion.ToString("dd/MM/yyyy")
        '    lcHojaExcel.Cells(2, 16).NumberFormat = "@"
        '    lcHojaExcel.Cells(2, 16).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        '    lcHojaExcel.Cells(2, 16).Value = lcFechaCreacion.ToString("hh:mm:ss tt")

        '    ' Recorrido de los datos de la consulta a la base de datos para introducir en la tabla del reporte

        '    ' Numero de la fila de los datos obtenidos de la consulta a la base de datos
        '    Dim lcFila As Integer
        '    ' Total de filas de los datos obtenidos de la consulta a la base de datos
        '    Dim lcTotalFilas As Integer = loDatosReporte.Tables(0).Rows.Count - 1
        '    ' Datos de la fila de la consulta a la base de datos
        '    Dim lcDatosFila As DataRow
        '    '' Control de agrupamiento 1 - (Clase de Articulo)(Cod_Cla)
        '    'Dim grupo1 As String = ""
        '    '' Control de agrupamiento 2 - (Codigo del Tipo de Documento)(Cod_Tip)
        '    'Dim grupo2 As String = ""
        '    ' Numero de la fila en el documento excel
        '    Dim lcNumFila As Integer = 7
        '    ' Numero de la fila en el documento excel, inicial para la sumatoria del total para el grupo de control 2
        '    Dim lcFilaIni As Integer = 7
        '    ' Numero de la fila en el documento excel, final para la sumatoria del total para el grupo de control 2
        '    Dim lcFilaFin As Integer = 0
        '    ' Construccion de la formula de total para el grupo de control 1
        '    Dim lcTotalDocumento As String = "= 0"

        '    ' Recorriendo las filas de los datos de la consulta a la base de datos
        '    For lcFila = 0 To lcTotalFilas
        '        ' Se extrae los datos de la fila
        '        lcDatosFila = loDatosReporte.Tables(0).Rows(lcFila)

        '        ' Se agrega la informacion en la tabla de los valores detallados

        '        lcHojaExcel.Cells(lcNumFila, 1).Value = lcDatosFila("Cod_Tip")
        '        lcHojaExcel.Cells(lcNumFila, 2).Value = lcDatosFila("Cod_Cli")
        '        lcHojaExcel.Cells(lcNumFila, 3).Value = lcDatosFila("Nom_Cli")
        '        lcHojaExcel.Cells(lcNumFila, 4).Value = lcDatosFila("Documento")
        '        lcHojaExcel.Cells(lcNumFila, 5).Value = lcDatosFila("Control")
        '        lcHojaExcel.Cells(lcNumFila, 6).Value = lcDatosFila("Fec_Ini")
        '        lcHojaExcel.Cells(lcNumFila, 7).Value = lcDatosFila("Fec_Fin")
        '        lcHojaExcel.Cells(lcNumFila, 8).Value = lcDatosFila("Dias")
        '        lcHojaExcel.Cells(lcNumFila, 9).Value = lcDatosFila("Fecha_Estimada")
        '        lcHojaExcel.Cells(lcNumFila, 10).Value = lcDatosFila("Cod_Ven")
        '        lcHojaExcel.Cells(lcNumFila, 11).Value = lcDatosFila("Cod_Mon")
        '        lcHojaExcel.Cells(lcNumFila, 12).Value = lcDatosFila("Mon_Bru")
        '        lcHojaExcel.Cells(lcNumFila, 13).Value = lcDatosFila("Mon_Imp1")
        '        lcHojaExcel.Cells(lcNumFila, 14).Value = lcDatosFila("Mon_Net")
        '        lcHojaExcel.Cells(lcNumFila, 15).Value = lcDatosFila("Mon_Sal")
        '        lcHojaExcel.Cells(lcNumFila, 16).Value = lcDatosFila("Accion")

        '        lcNumFila = lcNumFila + 1

        '    Next lcFila

        '    ' Se almacena el numero de la fila final del grupo de datos
        '    lcFilaFin = lcNumFila - 1
        '    ' Se coloca la etiqueta de total 
        '    lcRango = lcHojaExcel.Range("J" & CStr(lcNumFila) & ":K" & CStr(lcNumFila))
        '    lcRango.Select()
        '    lcRango.MergeCells = True
        '    lcRango.EntireRow.Font.Bold = True
        '    lcRango.Value = "Totales:"

        '    ' Se coloca la formula de total de todas las columnas

        '    lcHojaExcel.Cells(lcNumFila, 12).Formula = "=SUM(L" & CStr(lcFilaIni) & ":L" & CStr(lcNumFila - 1) & ")"
        '    lcHojaExcel.Cells(lcNumFila, 13).Formula = "=SUM(M" & CStr(lcFilaIni) & ":M" & CStr(lcNumFila - 1) & ")"
        '    lcHojaExcel.Cells(lcNumFila, 14).Formula = "=SUM(N" & CStr(lcFilaIni) & ":N" & CStr(lcNumFila - 1) & ")"
        '    lcHojaExcel.Cells(lcNumFila, 15).Formula = "=SUM(O" & CStr(lcFilaIni) & ":O" & CStr(lcNumFila - 1) & ")"

        '    lcNumFila = lcNumFila + 1

        '    ' Ajustamos el tamaño de las columnas
        '    lcRango = lcHojaExcel.Range("A1:A" & CStr(lcNumFila))
        '    lcRango.Select()
        '    lcRango.ColumnWidth = 6
        '    lcRango = lcHojaExcel.Range("B1:B" & CStr(lcNumFila))
        '    lcRango.Select()
        '    lcRango.ColumnWidth = 10
        '    lcRango = lcHojaExcel.Range("C1:C" & CStr(lcNumFila))
        '    lcRango.Select()
        '    lcRango.ColumnWidth = 35
        '    lcRango = lcHojaExcel.Range("D1:G" & CStr(lcNumFila))
        '    lcRango.Select()
        '    lcRango.ColumnWidth = 10
        '    lcRango = lcHojaExcel.Range("H1:H" & CStr(lcNumFila))
        '    lcRango.Select()
        '    lcRango.ColumnWidth = 7
        '    lcRango = lcHojaExcel.Range("I1:K" & CStr(lcNumFila))
        '    lcRango.Select()
        '    lcRango.ColumnWidth = 9
        '    lcRango = lcHojaExcel.Range("L1:O" & CStr(lcNumFila))
        '    lcRango.Select()
        '    lcRango.ColumnWidth = 10
        '    lcRango = lcHojaExcel.Range("P1:P" & CStr(lcNumFila))
        '    lcRango.Select()
        '    lcRango.ColumnWidth = 35
        '    lcRango.WrapText = True
        '    lcRango.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
        '    lcRango.Rows.AutoFit()
     

        '    ' Seleccionamos la primera celda del libro
        '    lcRango = lcHojaExcel.Range("A1")
        '    lcRango.Select()

        '    ' Cerramos y liberamos recursos
        '    mLiberar(lcRango)
        '    mLiberar(lcHojaExcel)
        '    'Guardamos los cambios del libro activo
        '    lcLibroExcel.Close(True, loFileName)

        '    mLiberar(lcLibroExcel)
        '    loAppExcel.Application.Quit()
        '    mLiberar(loAppExcel)

        'Catch loExcepcion As Exception
        '    Me.mEscribirConsulta(loExcepcion.Message)
        '    Me.Response.Flush()
        '    Me.Response.Close()

        '    Me.Response.End()

        'Finally
        '    ' Se forza el cierre del proceso excel
        '    'Dim Lista_Procesos() As Diagnostics.Process
        '    'Dim p As Diagnostics.Process
        '    'Lista_Procesos = Diagnostics.Process.GetProcessesByName("EXCEL")
        '    'For Each p In Lista_Procesos
        '    '    Try
        '    '        p.Kill()
        '    '    Catch
        '    '    End Try
        '    'Next
        '    GC.Collect()
        'End Try

    End Sub

    ''' <summary>
    ''' Cierre y liberacion de recursos de los objetos de la libreria Excel
    ''' </summary>
    ''' <param name="objeto"></param>
    ''' <remarks></remarks>
    Private Sub mLiberar(ByVal objeto As Object)
        'Try
        '    System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto)
        '    objeto = Nothing
        'Catch ex As Exception
        '    objeto = Nothing
        'End Try
    End Sub

End Class

'-------------------------------------------------------------------------------------------'
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' JJD: 22/09/08: Programacion inicial
'-------------------------------------------------------------------------------------------'
' CMS: 21/04/09: Estandarización del código y Corrección del estatus
'-------------------------------------------------------------------------------------------'
' CMS:  15/05/09: Filtro “Revisión:”
'-------------------------------------------------------------------------------------------'
' AAP:  30/06/09: Filtro “Sucursal:”
'-------------------------------------------------------------------------------------------'
' CMS:  03/07/09: Metodo de  Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS:  13/07/09: Se Agregaron los siguientes filtros: Zona, Tipo de Cliente,
'                 Clase de Cliente.
'                 Verificación de registros
'-------------------------------------------------------------------------------------------'
' CMS:  15/07/09: Multiplicación (*-1) al campo Mon_Net, Mon_Sal, Mon_Bru
'-------------------------------------------------------------------------------------------'
' RJG: 02/09/14: Se adaptó para generar la salida personalizada a Excel por medio de un     '
'                servicio externo (eFactory Servicios).                                     '
'-------------------------------------------------------------------------------------------'
