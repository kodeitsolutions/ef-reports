'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rImportacion_ValoresArticulos"
'-------------------------------------------------------------------------------------------'
Partial Class rImportacion_ValoresArticulos
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
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loSeleccion As New StringBuilder()


            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("SELECT      YEAR(Importaciones.Fec_Ini)                                     AS Ano, ")
            loSeleccion.AppendLine("            MONTH(Importaciones.Fec_Ini)                                    AS Mes, ")
            loSeleccion.AppendLine("            Importaciones.Fec_Ini                                           AS Fec_Ini,  ")
            loSeleccion.AppendLine("            Importaciones.Expediente                                        AS Expediente,  ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Tip_Ori                                 AS Tip_Ori, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Doc_Ori                                 AS Doc_Ori, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Ren_Ori                                 AS Ren_Ori,")
            loSeleccion.AppendLine("            Renglones_Importaciones.Cod_Art                                 AS Cod_Art, ")
            loSeleccion.AppendLine("            Articulos.Cod_Dep                                               AS Departamento, ")
            loSeleccion.AppendLine("            Articulos.Cod_Cla                                               AS Clase, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Can_Art1                                AS Can_Art1, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Fob                                 AS Mon_Fob, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Seg                                 AS Mon_Seg, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Fle                                 AS Mon_Fle, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Cif                                 AS Mon_Cif, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Alm                                 AS Mon_Alm, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Ipt                                 AS Mon_Ipt, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Por                                 AS Mon_Por, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Tra                                 AS Mon_Tra, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Per                                 AS Mon_Per, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Ban                                 AS Mon_Ban, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Adu                                 AS Mon_Adu, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Arc                                 AS Mon_Arc, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Ots1                                AS Mon_Ots1, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Ots2                                AS Mon_Ots2, ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Ots3                                AS Mon_Ots3,  ")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Ara                                 AS Mon_Ara,")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Gas_Fij                             AS Mon_Gas_Fij,")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Gas_Adi                             AS Mon_Gas_Adi,")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Com_Pag                             AS Mon_Com_Pag,")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Gas_Com                             AS Mon_Gas_Com,")
            loSeleccion.AppendLine("            Renglones_Importaciones.Mon_Net                                 AS Mon_Net,")
            loSeleccion.AppendLine("            ROUND(Renglones_Importaciones.Mon_Net/                          ")
            loSeleccion.AppendLine("                (CASE WHEN Renglones_Importaciones.Can_Art1>0 ")
            loSeleccion.AppendLine("                    THEN Renglones_Importaciones.Can_Art1 ")
            loSeleccion.AppendLine("                    ELSE 1 END), 2)                                         AS Mon_Net_Uni")
            loSeleccion.AppendLine("FROM        Importaciones ")
            loSeleccion.AppendLine("    JOIN    Renglones_Importaciones ")
            loSeleccion.AppendLine("        ON  (Renglones_Importaciones.Documento = Importaciones.Documento) ")
            loSeleccion.AppendLine("    JOIN    Articulos")
            loSeleccion.AppendLine("        ON  (Articulos.Cod_Art = Renglones_Importaciones.Cod_Art) ")
            loSeleccion.AppendLine(" WHERE     Importaciones.Documento BETWEEN " & lcParametro0Desde)
            loSeleccion.AppendLine(" 			AND " & lcParametro0Hasta)
            loSeleccion.AppendLine(" 			AND Importaciones.Fec_Ini BETWEEN " & lcParametro1Desde)
            loSeleccion.AppendLine(" 			AND " & lcParametro1Hasta)
            loSeleccion.AppendLine(" 			AND Importaciones.Cod_Pro BETWEEN " & lcParametro2Desde)
            loSeleccion.AppendLine(" 			AND " & lcParametro2Hasta)
            loSeleccion.AppendLine(" 			AND Renglones_Importaciones.Cod_Art BETWEEN " & lcParametro3Desde)
            loSeleccion.AppendLine(" 			AND " & lcParametro3Hasta)
            loSeleccion.AppendLine(" 			AND Importaciones.Cod_Mon BETWEEN " & lcParametro4Desde)
            loSeleccion.AppendLine(" 			AND " & lcParametro4Hasta)
            loSeleccion.AppendLine("ORDER BY      " & lcOrdenamiento & ", Importaciones.Fec_Ini, Importaciones.Fec_Fin")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")
            loSeleccion.AppendLine("")



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loSeleccion.ToString(), "curReportes")

                        
            '-------------------------------------------------------------------------------------------'
            ' Selección de opcion por excel (Microsoft Excel - xls)
            '-------------------------------------------------------------------------------------------'
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rImportacion_ValoresArticulos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrImportacion_ValoresArticulos.ReportSource = loObjetoReporte


        Catch loExcepcion As Exception

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")
        End Try

    End Sub

    Private Sub mGenerarArchivoExcel(ByVal loDatos As DataSet)

        '-------------------------------------------------------------------------------------------'
        ' Prepara los datos para enviarlos al servicio web de Excel.            '
        '-------------------------------------------------------------------------------------------'
        Dim loSalida As New IO.MemoryStream()
        loDatos.WriteXml(loSalida, XmlWriteMode.WriteSchema)

        '-------------------------------------------------------------------------------------------'
        ' Prepara los parámetros adicionales para enviarlos junto con los datos.'
        '-------------------------------------------------------------------------------------------'
        Dim lnDecimalesMonto As Integer = goOpciones.pnDecimalesParaMonto
        Dim lnDecimalesCantidad As Integer = goOpciones.pnDecimalesParaCantidad
        Dim lnDecimalesPorcentaje As Integer = goOpciones.pnDecimalesParaPorcentaje

        Dim loParametros As New NameValueCollection()
        loParametros.Add("lcNombreEmpresa", cusAplicacion.goEmpresa.pcNombre)
        loParametros.Add("lcRifEmpresa", cusAplicacion.goEmpresa.pcRifEmpresa)
        loParametros.Add("lnDecimalesMonto", lnDecimalesMonto.ToString())
        loParametros.Add("lnDecimalesCantidad", lnDecimalesCantidad.ToString())
        loParametros.Add("lnDecimalesPorcentaje", lnDecimalesPorcentaje.ToString())

        Dim loClienteWeb As New System.Net.WebClient()
        loClienteWeb.QueryString = loParametros

        '-------------------------------------------------------------------------------------------'
        ' Envía los datos y parámetros, y espera la respuesta.                  '
        '-------------------------------------------------------------------------------------------'
        Dim loRespuesta As Byte()
        Try
            Dim lcRuta As String = Me.MapPath("~\Framework\Xml\ParametrosGlobales.xml")
            Dim loParam As New System.Xml.XmlDocument()
            loParam.Load(lcRuta)
            Dim lcServicio As String = loParam.DocumentElement.GetAttribute("Servicios")

            loRespuesta = loClienteWeb.UploadData(lcServicio & "/Reportes/rImportacion_ValoresArticulos_xlsx.aspx", loSalida.GetBuffer())

        Catch ex As Exception
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
                                                                 ex.ToString(), vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return
        End Try

        '-------------------------------------------------------------------------------------------'
        ' Vemos si la respuesta es TextoPlano (error) o no (el archivo Excel    '
        ' generado). Si el tipo está vacio : error desconocido.                 '
        '-------------------------------------------------------------------------------------------'
        Dim loTipoRespuesta As String = loClienteWeb.ResponseHeaders("Content-Type")

        If String.IsNullOrEmpty(loTipoRespuesta) Then

            '-------------------------------------------------------------------------------------------'
            'Error no especificado!
            '-------------------------------------------------------------------------------------------'
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: El servicio que genera la salida XSLX no responde.<br/>", _
                                                                 vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return

        ElseIf loTipoRespuesta.ToLower().StartsWith("text/plain") Then

            Dim lcMensaje As String = UTF32Encoding.UTF8.GetString(loRespuesta)

            lcMensaje = Me.Server.HtmlEncode(lcMensaje)

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
                                                                 lcMensaje, vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return

        Else

            '-------------------------------------------------------------------------------------------'
            'Generación exitosa: la respuesta es el archivo en excel para descargar
            '-------------------------------------------------------------------------------------------'
            Me.Response.Clear()
            Me.Response.Buffer = True
            Me.Response.AppendHeader("content-disposition", "attachment; filename=rImportacion_ValoresArticulos.xlsx")
            Me.Response.ContentType = "application/excel"
            Me.Response.BinaryWrite(loRespuesta)
            Me.Response.End()

        End If


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
' RJG: 27/06/15: Programacion inicial.
'-------------------------------------------------------------------------------------------'
