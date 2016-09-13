'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rArchivoXML_RelacionRetencionISLR_RASA"
'-------------------------------------------------------------------------------------------'
Partial Class rArchivoXML_RelacionRetencionISLR_RASA
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcOrden As String = goReportes.pcOrden

        Dim ldFechaPeriodo As Date = cusAplicacion.goReportes.paParametrosIniciales(0)

        Dim ldInicioMes As Date = New Date(ldFechaPeriodo.Year, ldFechaPeriodo.Month, 1)
        Dim ldFinMes As Date = ldInicioMes.AddMonths(1).AddDays(-1)
        Dim lcFechaInicioSQL As String = goServicios.mObtenerCampoFormatoSQL(ldInicioMes, goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
        Dim lcFechaFinSQL As String = goServicios.mObtenerCampoFormatoSQL(ldFinMes, goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

        Try
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Pagos a Proveedores")
            loConsulta.AppendLine("SELECT       CAST(" & lcFechaInicioSQL & " AS DATETIME)      AS Periodo,     ")
            loConsulta.AppendLine("             '1'                                             AS Tipo,        ")
            loConsulta.AppendLine("				Proveedores.Cod_Pro					            AS Codigo,		")
            loConsulta.AppendLine("				Proveedores.Nom_Pro	                            AS Nombre,		")
            loConsulta.AppendLine("				Proveedores.Rif							        AS Rif,			")
            loConsulta.AppendLine("				Renglones_Pagos.Doc_Ori					        AS Factura,		")
            loConsulta.AppendLine("				Renglones_Pagos.Control					        AS Control,		")
            loConsulta.AppendLine("				Renglones_Pagos.Fec_Ini					        AS Emision,		")
            loConsulta.AppendLine("				Retenciones_Documentos.Cod_Ret					AS Concepto,	")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas			        AS Monto,		")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret			        AS Retenido,	")
            loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret			        AS Porcentaje 	")
            loConsulta.AppendLine("FROM			Cuentas_Pagar")
            loConsulta.AppendLine("		JOIN	Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.cla_des = Cuentas_Pagar.Cod_Tip")
            loConsulta.AppendLine("			AND Retenciones_Documentos.Clase = 'ISLR'")
            loConsulta.AppendLine("		JOIN	Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loConsulta.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("WHERE		Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")
            loConsulta.AppendLine("         AND Renglones_Pagos.Fec_Ini BETWEEN " & lcFechaInicioSQL)
            loConsulta.AppendLine("         AND " & lcFechaFinSQL)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UNION ALL		")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Órdenes de Pago a Proveedores")
            loConsulta.AppendLine("SELECT	    CAST(" & lcFechaInicioSQL & " AS DATETIME)      AS Periodo,     ")
            loConsulta.AppendLine("             '1'                                             AS Tipo,        ")
            loConsulta.AppendLine("				Proveedores.Cod_Pro					            AS Codigo,		")
            loConsulta.AppendLine("				Proveedores.Nom_Pro	                            AS Nombre,		")
            loConsulta.AppendLine("				Proveedores.Rif							        AS Rif,			")
            loConsulta.AppendLine("				Ordenes_Pagos.Documento						    AS Factura,     ")
            loConsulta.AppendLine("			    Ordenes_Pagos.Control						    AS Control,     ")
            loConsulta.AppendLine("			    Ordenes_Pagos.Fec_Ini						    AS Emision,     ")
            loConsulta.AppendLine("				Retenciones_Documentos.Cod_Ret					AS Concepto,    ")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas			        AS Monto,		")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret			        AS Retenido,	")
            loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret			        AS Porcentaje 	")
            loConsulta.AppendLine("FROM		Retenciones_Documentos")
            loConsulta.AppendLine("	JOIN	Ordenes_Pagos ON Ordenes_Pagos.Documento = Retenciones_Documentos.documento")
            loConsulta.AppendLine("	JOIN	Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            loConsulta.AppendLine("WHERE		Ordenes_Pagos.Status = 'Confirmado'")
            loConsulta.AppendLine("			AND	Retenciones_Documentos.Tip_Ori = 'Ordenes_Pagos'")
            loConsulta.AppendLine("			AND	Retenciones_Documentos.Clase = 'ISLR'")
            loConsulta.AppendLine("        AND  Ordenes_Pagos.Fec_Ini BETWEEN " & lcFechaInicioSQL)
            loConsulta.AppendLine("        AND " & lcFechaFinSQL)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UNION ALL		")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Retenciones sobre CxP a Proveedores")
            loConsulta.AppendLine("SELECT		CAST(" & lcFechaInicioSQL & " AS DATETIME)      AS Periodo,     ")
            loConsulta.AppendLine("             '1'                                             AS Tipo,        ")
            loConsulta.AppendLine("				Proveedores.Cod_Pro					            AS Codigo,		")
            loConsulta.AppendLine("				Proveedores.Nom_Pro	                            AS Nombre,		")
            loConsulta.AppendLine("				Proveedores.Rif							        AS Rif,			")
            loConsulta.AppendLine("				Documentos.Documento					        AS Factura,")
            loConsulta.AppendLine("				Documentos.Control						        AS Control,")
            loConsulta.AppendLine("				Documentos.Fec_Ini					            AS Emision,")
            loConsulta.AppendLine("				Retenciones_Documentos.Cod_Ret					AS Concepto,    ")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas			        AS Monto,		")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret			        AS Retenido,	")
            loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret			        AS Porcentaje 	")
            loConsulta.AppendLine("FROM			Cuentas_Pagar")
            loConsulta.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("WHERE		Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            loConsulta.AppendLine("        AND  Documentos.Fec_Ini BETWEEN " & lcFechaInicioSQL)
            loConsulta.AppendLine("        AND " & lcFechaFinSQL)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UNION ALL		")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Recibos de Nómina de Trabajadores ")
            loConsulta.AppendLine("SELECT      CAST(" & lcFechaInicioSQL & " AS DATETIME)      AS Periodo,")
            loConsulta.AppendLine("            '1'                                             AS Tipo,")
            loConsulta.AppendLine("            Trabajadores.Cod_Tra                            AS Codigo,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                            AS Nombre,")
            loConsulta.AppendLine("            Trabajadores.Rif                                AS Rif,")
            loConsulta.AppendLine("            CAST('0' AS CHAR(10))                           AS Factura,")
            loConsulta.AppendLine("            CAST('NA' AS CHAR(10))                          AS Control,")
            loConsulta.AppendLine("            Recibos.Fecha                                   AS Emision,")
            loConsulta.AppendLine("            CAST('001' AS CHAR(10))                         AS Concepto,")
            loConsulta.AppendLine("            Retenciones.Base                                AS Monto,")
            loConsulta.AppendLine("            Retenciones.Retenido                            AS Retenido,")
            loConsulta.AppendLine("            Retenciones.Porcentaje                          AS Porcentaje")
            loConsulta.AppendLine("FROM        Recibos")
            loConsulta.AppendLine("    JOIN    (   SELECT  Renglones_Recibos.Documento     AS Documento,")
            loConsulta.AppendLine("                        SUM(Renglones_Recibos.Mon_Net)  AS Retenido,")
            loConsulta.AppendLine("                        Renglones_Recibos.Val_Num       AS Porcentaje,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Renglones_Recibos.Val_Num>0")
            loConsulta.AppendLine("                            THEN (Renglones_Recibos.Mon_Net*100/Renglones_Recibos.Val_Num)")
            loConsulta.AppendLine("                            ELSE 0")
            loConsulta.AppendLine("                        end)                            AS Base")
            loConsulta.AppendLine("                FROM    Renglones_Recibos")
            loConsulta.AppendLine("                WHERE   Renglones_Recibos.Cod_Con IN ('R006', 'R305')")
            loConsulta.AppendLine("                GROUP BY Renglones_Recibos.Documento, ")
            loConsulta.AppendLine("                        Renglones_Recibos.Val_Num")
            loConsulta.AppendLine("        ) Retenciones")
            loConsulta.AppendLine("        ON  Retenciones.Documento = Recibos.Documento")
            loConsulta.AppendLine("    JOIN    Trabajadores ")
            loConsulta.AppendLine("        ON  Trabajadores.Cod_Tra = Recibos.Cod_Tra")
            loConsulta.AppendLine("        AND Trabajadores.Status = 'A'")
            loConsulta.AppendLine("WHERE       Recibos.Status = 'Confirmado'")
            loConsulta.AppendLine("        AND Recibos.Mon_Net > 0")
            loConsulta.AppendLine("        AND Recibos.Fecha BETWEEN " & lcFechaInicioSQL)
            loConsulta.AppendLine("        AND " & lcFechaFinSQL)
            loConsulta.AppendLine("UNION ALL		")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Recibos de Nómina de Trabajadores (ISLR) ")
            loConsulta.AppendLine("SELECT      CAST(" & lcFechaInicioSQL & " AS DATETIME)      AS Periodo,")
            loConsulta.AppendLine("            '1'                                             AS Tipo,")
            loConsulta.AppendLine("            Trabajadores.Cod_Tra                            AS Codigo,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                            AS Nombre,")
            loConsulta.AppendLine("            Trabajadores.Rif                                AS Rif,")
            loConsulta.AppendLine("            CAST('0' AS CHAR(10))                           AS Factura,")
            loConsulta.AppendLine("            CAST('NA' AS CHAR(10))                          AS Control,")
            loConsulta.AppendLine("            Recibos.Fecha                                   AS Emision,")
            loConsulta.AppendLine("            CAST('001' AS CHAR(10))                         AS Concepto,")
            loConsulta.AppendLine("            Retenciones.Porcentaje                          AS Monto,")
            loConsulta.AppendLine("            Retenciones.Retenido                            AS Retenido,")
            loConsulta.AppendLine("            Retenciones.Base                                AS Porcentaje")
            loConsulta.AppendLine("FROM        Recibos")
            loConsulta.AppendLine("    JOIN    (   SELECT  Renglones_Recibos.Documento     AS Documento,")
            loConsulta.AppendLine("                        SUM(Renglones_Recibos.Mon_Net)  AS Retenido,")
            loConsulta.AppendLine("                        Renglones_Recibos.Val_Num       AS Porcentaje,")
            loConsulta.AppendLine("                        SUM(CASE WHEN Renglones_Recibos.Val_Num>0")
            loConsulta.AppendLine("                            THEN (Renglones_Recibos.Mon_Net*100/Renglones_Recibos.Val_Num)")
            loConsulta.AppendLine("                            ELSE 0")
            loConsulta.AppendLine("                        end)                            AS Base")
            loConsulta.AppendLine("                FROM    Renglones_Recibos")
            loConsulta.AppendLine("                WHERE   Renglones_Recibos.Cod_Con IN ('R005')")
            loConsulta.AppendLine("                GROUP BY Renglones_Recibos.Documento, ")
            loConsulta.AppendLine("                        Renglones_Recibos.Val_Num")
            loConsulta.AppendLine("        ) Retenciones")
            loConsulta.AppendLine("        ON  Retenciones.Documento = Recibos.Documento")
            loConsulta.AppendLine("    JOIN    Trabajadores ")
            loConsulta.AppendLine("        ON  Trabajadores.Cod_Tra = Recibos.Cod_Tra")
            loConsulta.AppendLine("        AND Trabajadores.Status = 'A'")
            loConsulta.AppendLine("WHERE       Recibos.Status = 'Confirmado'")
            loConsulta.AppendLine("        AND Recibos.Mon_Net > 0")
            loConsulta.AppendLine("        AND Recibos.Fecha BETWEEN " & lcFechaInicioSQL)
            loConsulta.AppendLine("        AND " & lcFechaFinSQL)
            loConsulta.AppendLine("ORDER BY    Tipo, Codigo")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Dim lcSalida As String = Me.Request.QueryString("salida")
            If (lcSalida = "html") Then
                Me.mGenerarArchivoTxt(laDatosReporte)
                Return
            End If


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rArchivoXML_RelacionRetencionISLR_RASA", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrArchivoXML_RelacionRetencionISLR_RASA.ReportSource = loObjetoReporte

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

    Private Sub mGenerarArchivoTxt(ByVal laDatosReporte As DataSet)
        Dim loTabla As DataTable = laDatosReporte.Tables(0)
        Dim loLimpiaRIF As New Regex("[^A-Z0-9]", RegexOptions.Compiled)

        If (loTabla.Rows.Count = 0) Then
            'No se encontraron registros: dejar que el reporte salga normalmente
            Return
        End If


        Dim loRenglon As DataRow = loTabla.Rows(0)
        Dim ldFechaPeriodo As Date = CDate(loRenglon("Periodo"))
        Dim lcPeriodo As String = ldFechaPeriodo.ToString("yyyyMM")
        Dim lcNombreArchivo As String = "ISLR_" & ldFechaPeriodo.ToString("MMyy")

        Dim lcRifEmpresa As String = loLimpiaRIF.Replace(goEmpresa.pcRifEmpresa.ToUpper().Trim(), "")

        '**************************************************
        ' Creación del documento XML: declaración
        '**************************************************
        Dim loDocumento As New System.Xml.XmlDocument()

        Dim loDeclaracion As System.Xml.XmlDeclaration
        loDeclaracion = loDocumento.CreateXmlDeclaration("1.0", "utf-8", "")

        loDocumento.AppendChild(loDeclaracion)

        '**************************************************
        ' Creación del documento XML: Elemento Raiz (y datos de cabecera)
        '**************************************************
        Dim loRaiz As System.Xml.XmlElement
        Dim loAtributo As System.Xml.XmlAttribute

        loRaiz = loDocumento.CreateElement("RelacionRetencionesISLR")
        loDocumento.AppendChild(loRaiz)

        loAtributo = loDocumento.CreateAttribute("RifAgente")
        loAtributo.InnerText = lcRifEmpresa
        loRaiz.Attributes.Append(loAtributo)

        loAtributo = loDocumento.CreateAttribute("Periodo")
        loAtributo.InnerText = lcPeriodo
        loRaiz.Attributes.Append(loAtributo)


        '**************************************************
        ' Datos de trabajadores: montos a pagar
        '**************************************************
        Dim loRetencion As System.Xml.XmlElement
        Dim loElemento As System.Xml.XmlElement
        Dim lnCantidad As Integer = loTabla.Rows.Count
        For n As Integer = 0 To lnCantidad - 1
            loRenglon = loTabla.Rows(n)

            'Detalle de retención
            loRetencion = loDocumento.CreateElement("DetalleRetencion")
            loRaiz.AppendChild(loRetencion)

            'Datos de detalle: RIF 
            Dim lcRif As String = CStr(loRenglon("Rif")).ToUpper()
            lcRif = loLimpiaRIF.Replace(lcRif, "")

            loElemento = loDocumento.CreateElement("RifRetenido")
            loElemento.InnerText = lcRif
            loRetencion.AppendChild(loElemento)

            'Datos de detalle: Numero de Factura 
            loElemento = loDocumento.CreateElement("NumeroFactura")
            loElemento.InnerText = CStr(loRenglon("Factura")).Trim()
            loRetencion.AppendChild(loElemento)

            'Datos de detalle: Numero de Control 
            loElemento = loDocumento.CreateElement("NumeroControl")
            loElemento.InnerText = CStr(loRenglon("Control")).Trim()
            loRetencion.AppendChild(loElemento)

            'Datos de detalle: Fecha de la Operación
            Dim ldFecha As Date = CDate(loRenglon("Emision"))
            Dim lcFecha As String = ldFecha.ToString("dd/MM/yyyy")

            loElemento = loDocumento.CreateElement("FechaOperacion")
            loElemento.InnerText = lcFecha
            loRetencion.AppendChild(loElemento)

            'Datos de detalle: Concepto de Retención 
            loElemento = loDocumento.CreateElement("CodigoConcepto")
            loElemento.InnerText = CStr(loRenglon("Concepto")).Trim()
            loRetencion.AppendChild(loElemento)

            'Datos de detalle: Monto de la Operacion (2 decimales)
            Dim lnMonto As Decimal = Decimal.Round(CDec(loRenglon("Monto")), 2)
            Dim lcMonto As String = lnMonto.ToString("0.00")

            loElemento = loDocumento.CreateElement("MontoOperacion")
            loElemento.InnerText = lcMonto
            loRetencion.AppendChild(loElemento)

            'Datos de detalle: Porcentaje de Retencion (2 decimales)
            Dim lnPorcentaje As Decimal = Decimal.Round(CDec(loRenglon("Porcentaje")), 2)
            Dim lcPorcentaje As String = lnPorcentaje.ToString("0.00")

            loElemento = loDocumento.CreateElement("PorcentajeRetencion")
            loElemento.InnerText = lcPorcentaje
            loRetencion.AppendChild(loElemento)

        Next n

        Me.Response.Clear()
        Me.Response.Buffer = True
        Me.Response.AppendHeader("content-disposition", "attachment; filename=" & lcNombreArchivo & ".xml")
        Me.Response.ContentType = "text/xml"
        Me.Response.Write(loDocumento.InnerXml())
        Me.Response.End()

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' EAG: 06/10/15: Código Inicial.                                                            '
'-------------------------------------------------------------------------------------------'
