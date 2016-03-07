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
' Inicio de clase "rResumen_Gerencial"
'-------------------------------------------------------------------------------------------'
Partial Class rResumen_Gerencial
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT	Comprobantes.Documento, ")
            loComandoSeleccionar.AppendLine("           YEAR(Comprobantes.Fec_Ini)  AS  Anno, ")
            loComandoSeleccionar.AppendLine("           MONTH(Comprobantes.Fec_Ini) AS  Mes, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Resumen, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Tipo, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Origen, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Integracion, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Status, ")
            loComandoSeleccionar.AppendLine("           Comprobantes.Notas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Renglon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Cue, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Cen, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Gas, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Act, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Tip, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Cla, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Tasa, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Mon_Deb, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Mon_Hab, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Comentario, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes.Cod_Reg ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes01 ")
            loComandoSeleccionar.AppendLine(" FROM      Comprobantes, ")
            loComandoSeleccionar.AppendLine("           Renglones_Comprobantes ")
            loComandoSeleccionar.AppendLine(" WHERE     Comprobantes.Documento                      =   Renglones_Comprobantes.Documento ")
            loComandoSeleccionar.AppendLine("           And Comprobantes.Documento                  Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           And Comprobantes.Fec_Ini                    Between " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           And Renglones_Comprobantes.Cod_Mon          Between " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("           And " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           And YEAR(Renglones_Comprobantes.Fec_Ini)    =   YEAR(Comprobantes.Fec_Ini) ")
            loComandoSeleccionar.AppendLine("           And MONTH(Renglones_Comprobantes.Fec_Ini)   =   MONTH(Comprobantes.Fec_Ini) ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes01.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Nom_Cue,1,30) END) AS Nom_Cue, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Nom_Alt,1,30) END) AS Nom_Alt, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Cla_Alt,1,30) END) AS Cla_Alt, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Gru_Alt,1,30) END) AS Gru_Alt, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Seg_Alt,1,30) END) AS Seg_Alt, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Fam_Alt,1,30) END) AS Fam_Alt, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Caracter1,1,30) END) AS Caracter1, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Caracter2,1,30) END) AS Caracter2, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Caracter3,1,30) END) AS Caracter3, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Caracter4,1,30) END) AS Caracter4, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes01.Cod_Cue = '' THEN '' ELSE SUBSTRING(Cuentas_Contables.Caracter5,1,30) END) AS Caracter5 ")

            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes02 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes01 LEFT JOIN Cuentas_Contables ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes01.Cod_Cue   =   Cuentas_Contables.Cod_Cue ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes02.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes02.Cod_Cen = '' THEN '' ELSE SUBSTRING(Centros_Costos.Nom_Cen,1,30) END)         AS Nom_Cen, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes02.Cod_Cen = '' THEN '' ELSE SUBSTRING(Centros_Costos.Recurso,1,30) END)         AS CenCos_Recurso, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes02.Cod_Cen = '' THEN '' ELSE SUBSTRING(Centros_Costos.Categoria,1,30) END)       AS CenCos_Categoria, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes02.Cod_Cen = '' THEN '' ELSE SUBSTRING(Centros_Costos.Departamento,1,30) END)    AS CenCos_Departamento, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes02.Cod_Cen = '' THEN '' ELSE SUBSTRING(Centros_Costos.Seccion,1,30) END)         AS CenCos_Seccion ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes03 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes02 LEFT JOIN Centros_Costos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes02.Cod_Cen   =   Centros_Costos.Cod_Cen ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes03.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes03.Cod_Gas = '' THEN '' ELSE SUBSTRING(Cuentas_Gastos.Nom_Gas,1,30) END) AS Nom_Gas ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes04 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes03 LEFT JOIN Cuentas_Gastos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes03.Cod_Gas   =   Cuentas_Gastos.Cod_Gas ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes04.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes04.Cod_Act = '' THEN '' ELSE SUBSTRING(Activos_Fijos.Nom_Act,1,30) END) AS Nom_Act ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes05 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes04 LEFT JOIN Activos_Fijos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes04.Cod_Act   =   Activos_Fijos.Cod_Act ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes05.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes05.Cod_Tip = '' THEN '' ELSE SUBSTRING(Tipos_Documentos.Nom_Tip,1,30) END) AS Nom_Tip ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes06 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes05 LEFT JOIN Tipos_Documentos ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes05.Cod_Tip   =   Tipos_Documentos.Cod_Tip ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes06.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes06.Cod_Cla = '' THEN '' ELSE SUBSTRING(Clasificadores.Nom_Cla,1,30) END) AS Nom_Cla ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes07 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes06 LEFT JOIN Clasificadores ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes06.Cod_Cla   =   Clasificadores.Cod_Cla ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes07.*, ")
            loComandoSeleccionar.AppendLine("           (CASE WHEN #tmpComprobantes07.Cod_Mon = '' THEN '' ELSE SUBSTRING(Monedas.Nom_Mon,1,30) END) AS Nom_Mon ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes08 ")
            loComandoSeleccionar.AppendLine(" FROM      #tmpComprobantes07 LEFT JOIN Monedas ")
            loComandoSeleccionar.AppendLine("           ON #tmpComprobantes07.Cod_Mon   =   Monedas.Cod_Mon ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes08.*, ")
            loComandoSeleccionar.AppendLine(" 			SUBSTRING(#tmpComprobantes08.COD_CUE,1,4) AS SsCpte1, ")
            loComandoSeleccionar.AppendLine(" 			SUBSTRING(#tmpComprobantes08.COD_CUE,1,2) AS SsCpte2, ")
            loComandoSeleccionar.AppendLine(" 			SUBSTRING(#tmpComprobantes08.COD_CUE,1,1) AS Balance ")
            loComandoSeleccionar.AppendLine(" INTO      #tmpComprobantes09 ")
            loComandoSeleccionar.AppendLine(" FROM		#tmpComprobantes08 ")

            loComandoSeleccionar.AppendLine(" SELECT	Cod_Cue, Nom_Cue ")
            loComandoSeleccionar.AppendLine(" INTO		#tmpCuentasContables ")
            loComandoSeleccionar.AppendLine(" FROM		Cuentas_Contables ")
            loComandoSeleccionar.AppendLine(" WHERE		LEN(Cod_Cue) = '1' ")

            loComandoSeleccionar.AppendLine(" SELECT	#tmpComprobantes09.*, ")
            loComandoSeleccionar.AppendLine("           #tmpCuentasContables.Nom_Cue AS Desc_Bal ")
            loComandoSeleccionar.AppendLine(" FROM		#tmpComprobantes09, #tmpCuentasContables ")
            loComandoSeleccionar.AppendLine(" WHERE		#tmpComprobantes09.Balance = #tmpCuentasContables.Cod_Cue ")
            loComandoSeleccionar.AppendLine(" ORDER BY	#tmpComprobantes09.Documento, #tmpComprobantes09.Renglon ASC ")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------'
            ' Selección de opcion por excel (Microsoft Excel - xls)
            '-------------------------------------------------------------------------------------------'
            If (Me.Request.QueryString("salida").ToLower() = "xls") Then
                ' Genera el archivo a partir de la tabla de datos y termina la ejecución
                Me.mGenerarArchivoExcel(laDatosReporte)

            End If

            '-------------------------------------------------------------------------------------------'
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------'
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If





            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rResumen_Gerencial", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrResumen_Gerencial.ReportSource = loObjetoReporte

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
        ' Prepara los datos para enviarlos al servicio web de Excel.
        '-------------------------------------------------------------------------------------------'
        Dim loSalida As New IO.MemoryStream()
        loDatos.WriteXml(loSalida, XmlWriteMode.WriteSchema)


        '-------------------------------------------------------------------------------------------'
        ' Prepara los parámetros adicionales para enviarlos junto con los datos.
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

        Dim loClienteWeb As New WebClient()
        loClienteWeb.QueryString = loParametros

        '-------------------------------------------------------------------------------------------'
        ' Envía los datos y parámetros, y espera la respuesta.
        '-------------------------------------------------------------------------------------------'
        Dim loRespuesta As Byte()
        Try
            Dim lcRuta As String = Me.MapPath("~\Framework\Xml\ParametrosGlobales.xml")
            Dim loParam As New System.Xml.XmlDocument()
            loParam.Load(lcRuta)
            Dim lcServicio As String = loParam.DocumentElement.GetAttribute("Servicios")

            loRespuesta = loClienteWeb.UploadData(lcServicio & "/Reportes/rResumen_Gerencial_xlsx.aspx", loSalida.GetBuffer())
        Catch ex As Exception
            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Proceso no Completado", _
                                                                 "No fue posible generar el reporte solicitado. Información Adicional: <br/>" & _
                                                                 ex.ToString(), vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, "600px", "500px")
            Return
        End Try

        '-------------------------------------------------------------------------------------------'
        ' Vemos si la respuesta es TextoPlano (error) o no (el archivo Excel
        ' generado). Si el tipo está vacio : error desconocido.
        '-------------------------------------------------------------------------------------------'
        Dim loTipoRespuesta As String = loClienteWeb.ResponseHeaders("Content-Type")

        If String.IsNullOrEmpty(loTipoRespuesta) Then
            'Error no especificado!
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
            Me.Response.AppendHeader("content-disposition", "attachment; filename=rResumen_Gerencial_xlsx.xlsx")
            Me.Response.ContentType = "application/excel"
            Me.Response.BinaryWrite(loRespuesta)
            Me.Response.End()

        End If


    End Sub

    Private Sub mLiberar(ByVal objeto As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objeto)
            objeto = Nothing
        Catch ex As Exception
            objeto = Nothing
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
' JJD: 23/02/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 17/08/09: Metodo de ordenamiento
'-------------------------------------------------------------------------------------------'
' MAT: 16/05/11: Mejora de la vista de Diseño
'-------------------------------------------------------------------------------------------'
' JJD: 18/11/14: Inclusion del envio a Excell segun el nuevo esquema de trabajo
'-------------------------------------------------------------------------------------------'