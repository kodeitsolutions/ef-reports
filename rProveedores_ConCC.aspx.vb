'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System
Imports System.Data
Imports System.Collections.Specialized
Imports System.Net
Imports cusAplicacion

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rProveedores_ConCC"
'-------------------------------------------------------------------------------------------'
Partial Class rProveedores_ConCC
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
            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Tabla temporal con los registros a listar")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpRegistros( Codigo VARCHAR(30) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Nombre VARCHAR(100) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Estatus VARCHAR(15) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Contable XML")
            loComandoSeleccionar.AppendLine("                            );")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpRegistros(Codigo, Nombre, Estatus, Contable)")
            loComandoSeleccionar.AppendLine("SELECT  Cod_Pro, ")
            loComandoSeleccionar.AppendLine("        Nom_Pro,")
            loComandoSeleccionar.AppendLine("        (CASE Status ")
            loComandoSeleccionar.AppendLine("            WHEN 'A' THEN 'Activo'  ")
            loComandoSeleccionar.AppendLine("            WHEN 'I' THEN 'Inactivo'")
            loComandoSeleccionar.AppendLine("            ELSE 'Suspendido'  ")
            loComandoSeleccionar.AppendLine("        END) AS Status,")
            loComandoSeleccionar.AppendLine("        Contable")
            loComandoSeleccionar.AppendLine("FROM   Proveedores")
            loComandoSeleccionar.AppendLine("WHERE	Proveedores.Cod_Pro     Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Proveedores.Status IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine(" 		AND Proveedores.Cod_Ven Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Proveedores.Cod_Zon Between " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Proveedores.Cod_Pai Between " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Proveedores.Cod_Cla Between " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Proveedores.Cod_Tip Between " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine(" 		AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine(" 		AND Proveedores.Status  =   'A' ")

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- En el SELECT final se expande el XML Contable para obtener las ")
            loComandoSeleccionar.AppendLine("-- Cuentas Contables, de Gastos y Centros de Costos de cada página del registro ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  CASE WHEN (LEN(Detalles.Cue_Con_Codigo) > '1' AND (LEN(Detalles.Cue_Con_Codigo) < '9' OR LEN(Detalles.Cue_Con_Codigo) > '9')) THEN '******' ELSE '' END	AS Asteriscos,")
            loComandoSeleccionar.AppendLine("        #tmpRegistros.Codigo                                AS Codigo,")
            loComandoSeleccionar.AppendLine("        #tmpRegistros.Nombre                                AS Nombre,")
            loComandoSeleccionar.AppendLine("        #tmpRegistros.Estatus                               AS Estatus,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Numero, 1)                        AS Numero,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Pagina, '')                       AS Pagina,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cue_Con_Codigo, '')               AS Cue_Con_Codigo,")
            loComandoSeleccionar.AppendLine("        COALESCE(Cuentas_Contables.Nom_Cue, '')             AS Cue_Con_Nombre,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cue_Gas_Codigo, '')               AS Cue_Gas_Codigo,")
            loComandoSeleccionar.AppendLine("        COALESCE(Cuentas_Gastos.Nom_Gas, '')                AS Cue_Gas_Nombre,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cen_Cos_Codigo, '')               AS Cen_Cos_Codigo,")
            loComandoSeleccionar.AppendLine("        COALESCE(Centros_Costos.Nom_Cen, '')                AS Cen_Cos_Nombre,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cen_Cos_Porcentaje, 0)            AS Cen_Cos_Porcentaje ")
            loComandoSeleccionar.AppendLine("FROM    #tmpRegistros")
            loComandoSeleccionar.AppendLine("    LEFT JOIN ( SELECT  Codigo,")
            loComandoSeleccionar.AppendLine("                        (Ficha.C.value('@n[1]', 'VARCHAR(MAX)')+1) AS Numero,")
            loComandoSeleccionar.AppendLine("                        Ficha.C.value('@nombre[1]', 'VARCHAR(MAX)') AS Pagina,")
            loComandoSeleccionar.AppendLine("                        Ficha.C.value('./cue_con[1]', 'VARCHAR(MAX)') AS Cue_Con_Codigo,")
            loComandoSeleccionar.AppendLine("                        Ficha.C.value('./cue_gas[1]', 'VARCHAR(MAX)') AS Cue_Gas_Codigo,")
            loComandoSeleccionar.AppendLine("                        Costos.C.value('@codigo[1]', 'VARCHAR(MAX)') AS Cen_Cos_Codigo,")
            loComandoSeleccionar.AppendLine("                        CAST(Costos.C.value('@porcentaje[1]', 'VARCHAR(MAX)') AS DECIMAL(28,10)) AS Cen_Cos_Porcentaje")
            loComandoSeleccionar.AppendLine("                FROM    #tmpRegistros")
            loComandoSeleccionar.AppendLine("                    CROSS APPLY Contable.nodes('contable/ficha') AS Ficha(C)")
            loComandoSeleccionar.AppendLine("                    OUTER APPLY Contable.nodes('contable/ficha/centro_costo') AS Costos(C)")
            loComandoSeleccionar.AppendLine("            ) Detalles")
            loComandoSeleccionar.AppendLine("        ON  Detalles.Codigo = #tmpRegistros.Codigo")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Cuentas_Contables")
            loComandoSeleccionar.AppendLine("        ON Cuentas_Contables.Cod_Cue = Detalles.Cue_Con_Codigo")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Cuentas_Gastos")
            loComandoSeleccionar.AppendLine("        ON Cuentas_Gastos.Cod_Gas = Detalles.Cue_Gas_Codigo")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Centros_Costos")
            loComandoSeleccionar.AppendLine("        ON Centros_Costos.Cod_Cen = Detalles.Cen_Cos_Codigo")
            loComandoSeleccionar.AppendLine("ORDER BY #tmpRegistros.Codigo, COALESCE(Detalles.Numero, 1)")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            ''-------------------------------------------------------------------------------------------'
            '' Selección de opcion por excel (Microsoft Excel - xls)                                     '
            ''-------------------------------------------------------------------------------------------'
            'If (Me.Request.QueryString("salida").ToLower() = "xls") Then
            '    ' Genera el archivo a partir de la tabla de datos y termina la ejecución
            '    Me.mGenerarArchivoExcel(laDatosReporte)

            'End If

            '-------------------------------------------------------------------------------------------'
            ' Verificando si el select (tabla nº0) trae registros                                       '
            '-------------------------------------------------------------------------------------------'
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProveedores_ConCC", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrProveedores_ConCC.ReportSource = loObjetoReporte


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


    'Private Sub mGenerarArchivoExcel(ByVal loDatos As DataSet)

    '    '-------------------------------------------------------------------------------------------'
    '    ' Prepara los datos para enviarlos al servicio web de Excel.                                '
    '    '-------------------------------------------------------------------------------------------'
    '    Dim loSalida As New IO.MemoryStream()
    '    loDatos.WriteXml(loSalida, XmlWriteMode.WriteSchema)

    '    '-------------------------------------------------------------------------------------------'
    '    ' Prepara los parámetros adicionales para enviarlos junto con los datos.                    '
    '    '-------------------------------------------------------------------------------------------'
    '    Dim lnDecimalesMonto As Integer = goOpciones.pnDecimalesParaMonto
    '    Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
    '    Dim lnDecimalesPorcentaje As Integer = goOpciones.pnDecimalesParaPorcentaje

    '    Dim loParametros As New NameValueCollection()
    '    loParametros.Add("lcNombreEmpresa", cusAplicacion.goEmpresa.pcNombre)
    '    loParametros.Add("lcRifEmpresa", cusAplicacion.goEmpresa.pcRifEmpresa)
    '    loParametros.Add("lnDecimalesMonto", lnDecimalesMonto.ToString())
    '    loParametros.Add("lnDecimalesCantidad", lnDecimalesCantidad.ToString())
    '    loParametros.Add("lnDecimalesPorcentaje", lnDecimalesPorcentaje.ToString())

    '    Dim loClienteWeb As New WebClient()
    '    loClienteWeb.QueryString = loParametros

    '    '-------------------------------------------------------------------------------------------'
    '    ' Envía los datos y parámetros, y espera la respuesta.                                      '
    '    '-------------------------------------------------------------------------------------------'
    '    Dim loRespuesta As Byte()
    '    Try
    '        Dim lcRuta As String = Me.MapPath("~\Framework\Xml\ParametrosGlobales.xml")
    '        Dim loParam As New System.Xml.XmlDocument()
    '        loParam.Load(lcRuta)
    '        Dim lcServicio As String = loParam.DocumentElement.GetAttribute("Servicios")

    '        loRespuesta = loClienteWeb.UploadData(lcServicio & "/Reportes/rProveedores_ConCC_xlsx.aspx", loSalida.GetBuffer())
    '    Catch ex As Exception
    '        Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
    '                                                             "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
    '                                                             ex.ToString(), vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
    '        Return
    '    End Try

    '    '-------------------------------------------------------------------------------------------'
    '    ' Vemos si la respuesta es TextoPlano (error) o no (el archivo Excel                        '
    '    ' generado). Si el tipo está vacio : error desconocido.                                     '
    '    '-------------------------------------------------------------------------------------------'
    '    Dim loTipoRespuesta As String = loClienteWeb.ResponseHeaders("Content-Type")

    '    If String.IsNullOrEmpty(loTipoRespuesta) Then
    '        'Error no especificado!
    '        Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
    '                                                             "No fue posible generar el reporte solicitado. Información Adicional: El servicio que genera la salida XSLX no responde.<br/>", _
    '                                                             vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
    '        Return

    '    ElseIf loTipoRespuesta.ToLower().StartsWith("text/plain") Then

    '        Dim lcMensaje As String = UTF32Encoding.UTF8.GetString(loRespuesta)

    '        lcMensaje = Me.Server.HtmlEncode(lcMensaje)

    '        Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
    '                                                             "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
    '                                                             lcMensaje, vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
    '        Return

    '    Else
    '        'Generación exitosa: la respuesta es el archivo en excel para descargar

    '        Me.Response.Clear()
    '        Me.Response.Buffer = True
    '        Me.Response.AppendHeader("content-disposition", "attachment; filename=rProveedores_ConCC_xlsx.xlsx")
    '        Me.Response.ContentType = "application/excel"
    '        Me.Response.BinaryWrite(loRespuesta)
    '        Me.Response.End()

    '    End If

    'End Sub

End Class

'-------------------------------------------------------------------------------------------'
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' JJD: 05/11/14: Codigo inicial. Adecuacion para mostrar las cuentas contables              '
'-------------------------------------------------------------------------------------------'
' JJD: 18/12/14: Inclusion del Len de la Cuenta Contable                                    '
'-------------------------------------------------------------------------------------------'
