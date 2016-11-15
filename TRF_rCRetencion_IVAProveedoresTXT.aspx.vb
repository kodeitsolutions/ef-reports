'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rCRetencion_IVAProveedoresTXT"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rCRetencion_IVAProveedoresTXT
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcAño As String = CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Year)
            Dim lcMes As String = CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Month)
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("               Cuentas_Pagar.Documento             AS Documento,")
            loComandoSeleccionar.AppendLine("               Cuentas_Pagar.Doc_Ori               AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Num_Com		AS Numero_Comprobante,			")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Documento				AS Numero_DocumentoRet,			")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Control				AS Numero_ControlRet,			")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("			    Documentos.Factura				    AS Doc_Pro002,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Control					AS Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Mon_Net				AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Exe					AS Monto_ExentoDoc,				")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Bas1					AS Monto_BaseDoc,				")
            loComandoSeleccionar.AppendLine("				Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,		")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Imp1					AS Monto_ImpuestoDoc,			")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Proveedores.Dir_Fis					AS Direccion")
            'loComandoSeleccionar.AppendLine("				" & lcAño & "						AS Anio,")
            'loComandoSeleccionar.AppendLine("				" & lcMes & "						AS Mes")
            loComandoSeleccionar.AppendLine("INTO			#tabRetenciones")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("		JOIN	Pagos")
            loComandoSeleccionar.AppendLine("			ON	Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ")
            loComandoSeleccionar.AppendLine("			ON	Retenciones_Documentos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.clase	= 'IMPUESTO'")
            loComandoSeleccionar.AppendLine("		JOIN	Renglones_Pagos ")
            loComandoSeleccionar.AppendLine("			ON	Renglones_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("		LEFT JOIN	Cuentas_Pagar AS Documentos									")
            loComandoSeleccionar.AppendLine("			ON	Documentos.Documento = Renglones_Pagos.Doc_Ori					")
            loComandoSeleccionar.AppendLine("			AND	Documentos.Cod_Tip = Renglones_Pagos.Cod_Tip 					")
            'loComandoSeleccionar.AppendLine("		JOIN        Cuentas_Pagar                       AS  Documentos002 ")
            'loComandoSeleccionar.AppendLine("			ON          Documentos.Doc_Ori                =   Documentos002.Documento")
            loComandoSeleccionar.AppendLine("		JOIN	Proveedores")
            loComandoSeleccionar.AppendLine("			ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("		LEFT JOIN Retenciones")
            loComandoSeleccionar.AppendLine("			ON	Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")

            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		")

            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("               Cuentas_Pagar.Documento             AS Documento,")
            loComandoSeleccionar.AppendLine("               Cuentas_Pagar.Doc_Ori               AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Num_Com		AS Numero_Comprobante,				")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Documento				AS Numero_DocumentoRet,				")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Control				AS Numero_ControlRet,				")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("			    Documentos.Factura				    AS Doc_Pro002,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Control					AS Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Net					AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Exe					AS Monto_ExentoDoc,					")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Bas1					AS Monto_BaseDoc,					")
            loComandoSeleccionar.AppendLine("				Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,			")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Imp1					AS Monto_ImpuestoDoc,				")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Proveedores.Dir_Fis					AS Direccion")
            'loComandoSeleccionar.AppendLine("				" & lcAño & "						AS Anio,")
            'loComandoSeleccionar.AppendLine("				" & lcMes & "						AS Mes")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ")
            loComandoSeleccionar.AppendLine("			ON	Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            'loComandoSeleccionar.AppendLine("		    JOIN        Cuentas_Pagar                       AS  Documentos002 ")
            'loComandoSeleccionar.AppendLine("			ON          Documentos.Doc_Ori                =   Documentos002.Documento")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos")
            loComandoSeleccionar.AppendLine("			ON	Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Proveedores ")
            loComandoSeleccionar.AppendLine("			ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("		LEFT JOIN	Retenciones")
            loComandoSeleccionar.AppendLine("			ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")

            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro0Hasta)

            'loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento & ", Fecha_Retencion ASC")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("UPDATE #tabRetenciones")
            loComandoSeleccionar.AppendLine("SET #tabRetenciones.Cod_Pro = (SELECT CASE WHEN (ISNULL((SELECT Cod_Ter")
            loComandoSeleccionar.AppendLine("                                                      FROM Retenciones_Renglones")
            loComandoSeleccionar.AppendLine("                                                      WHERE Doc_Des = #tabRetenciones.Documento   ")
            loComandoSeleccionar.AppendLine("                                                      AND Doc_Ori = #tabRetenciones.Doc_Ori),'') = '')  ")
            loComandoSeleccionar.AppendLine("                        THEN #tabRetenciones.Cod_Pro")
            loComandoSeleccionar.AppendLine("                        ELSE (SELECT Cod_Ter")
            loComandoSeleccionar.AppendLine("                              FROM Retenciones_Renglones")
            loComandoSeleccionar.AppendLine("                              WHERE Doc_Des = #tabRetenciones.Documento   ")
            loComandoSeleccionar.AppendLine("                              AND Doc_Ori = #tabRetenciones.Doc_Ori)")
            loComandoSeleccionar.AppendLine("                        END),")
            loComandoSeleccionar.AppendLine("      #tabRetenciones.Rif = (SELECT CASE WHEN (ISNULL((SELECT Cod_Ter")
            loComandoSeleccionar.AppendLine("                                                      FROM Retenciones_Renglones")
            loComandoSeleccionar.AppendLine("                                                      WHERE Doc_Des = #tabRetenciones.Documento   ")
            loComandoSeleccionar.AppendLine("                                                      AND Doc_Ori = #tabRetenciones.Doc_Ori),'') <>'')  ")
            loComandoSeleccionar.AppendLine("                                    THEN (SELECT Rif")
            loComandoSeleccionar.AppendLine("                                          FROM Proveedores")
            loComandoSeleccionar.AppendLine("                                          WHERE Cod_Pro = (SELECT Cod_Ter")
            loComandoSeleccionar.AppendLine("                                                                  FROM Retenciones_Renglones")
            loComandoSeleccionar.AppendLine("                                                                  WHERE Doc_Des = #tabRetenciones.Documento")
            loComandoSeleccionar.AppendLine("                                                                  AND Doc_Ori = #tabRetenciones.Doc_Ori))")
            loComandoSeleccionar.AppendLine("                                    ELSE #tabRetenciones.Rif END")
            loComandoSeleccionar.AppendLine("                              )")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT		") ' & lcRif    & "					AS Rif_Empresa,")
            'loComandoSeleccionar.AppendLine("			(Anio + Mes)						AS Periodo,")
            loComandoSeleccionar.AppendLine("			Fecha_Retencion						AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("			'C'									AS Tipo_Operacion,")
            loComandoSeleccionar.AppendLine("			CASE Tipo_Documento")
            loComandoSeleccionar.AppendLine("				WHEN 'FACT' THEN '01'")
            loComandoSeleccionar.AppendLine("				WHEN 'N/DB' THEN '02'")
            loComandoSeleccionar.AppendLine("				WHEN 'N/CR' THEN '03'")
            loComandoSeleccionar.AppendLine("			END									AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("			Rif                             	AS Rif_Proveedor,")
            loComandoSeleccionar.AppendLine("			Numero_Documento					AS Documento_Origen,")
            loComandoSeleccionar.AppendLine("			Doc_Pro002,")
            loComandoSeleccionar.AppendLine("			Numero_ControlDoc					AS Control_Origen, ")
            loComandoSeleccionar.AppendLine("			Monto_Documento						AS Monto_Neto,")
            loComandoSeleccionar.AppendLine("			Monto_BaseDoc					    AS Base_Imponible,")
            loComandoSeleccionar.AppendLine("			Monto_Retenido						AS Monto_Retencion,")
            loComandoSeleccionar.AppendLine("			'0'									AS Numero_Factura,")
            loComandoSeleccionar.AppendLine("			Numero_Comprobante					AS Numero_Comprobante,")
            loComandoSeleccionar.AppendLine("			Monto_ExentoDoc						AS Monto_Exento,")
            loComandoSeleccionar.AppendLine("			Porcentaje_ImpuestoDoc				AS Porcentaje_Impuesto,")
            loComandoSeleccionar.AppendLine("			'0'									AS Numero_Expediente")
            loComandoSeleccionar.AppendLine("FROM		#tabRetenciones")
            loComandoSeleccionar.AppendLine("ORDER BY 	Fecha_Retencion, Tipo_Documento")
            loComandoSeleccionar.AppendLine("	")
            loComandoSeleccionar.AppendLine("DROP TABLE #tabRetenciones")
            loComandoSeleccionar.AppendLine("")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")


            '-------------------------------------------------------------------------------------------------------
            ' Genera el archivo de texto    
            '-------------------------------------------------------------------------------------------------------
            Dim loLimpiarRIF As New Regex("[^a-zA-Z0-9]")

            Dim lcRif As String = loLimpiarRIF.Replace(goEmpresa.pcRifEmpresa, "")
            Dim lcPeriodo As String = Strings.Format(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)), "yyyyMM")
            Dim loSalida As New StringBuilder()

            If laDatosReporte.Tables(0).Rows.Count <= 0 Then

                loSalida.Append(lcRif).Append(vbTab)
                loSalida.Append(lcPeriodo).Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.Append("0").Append(vbTab)
                loSalida.AppendLine()

            Else
                For Each loFila As DataRow In laDatosReporte.Tables(0).Rows

                    Dim ldFecha_Retencion As Date = CDate(loFila("Fecha_Retencion"))
                    Dim lcTipo_Operacion As String = CStr(loFila("Tipo_Operacion")).Trim()
                    Dim lcTipo_Documento As String = CStr(loFila("Tipo_Documento")).Trim()

                    Dim lcRif_Proveedor As String = loLimpiarRIF.Replace(CStr(loFila("Rif_Proveedor")), "")

                    Dim lcDocumento_Origen As String = CStr(loFila("Doc_Pro002")).Trim()
                    'Dim lcDocumento_Origen As String = CStr(loFila("Documento_Origen")).Trim()
                    Dim lcControl_Origen As String = CStr(loFila("Control_Origen")).Trim()
                    Dim lnMonto_Neto As Decimal = CDec(loFila("Monto_Neto"))
                    Dim lnBase_Imponible As Decimal = CDec(loFila("Base_Imponible"))
                    Dim lnMonto_Retencion As Decimal = CDec(loFila("Monto_Retencion"))
                    Dim lcNumero_Factura As String = CStr(loFila("Numero_Factura")).Trim()

                    Dim lcNumero_Comprobante As String = Strings.Format(lcPeriodo & CStr(loFila("Numero_Comprobante")).Trim())

                    'Dim lcNumero_Comprobante As String = Strings.Format(CStr(loFila("Numero_Comprobante")).Trim(), "00000000000000")
                    Dim lnMonto_Exento As Decimal = CDec(loFila("Monto_Exento"))
                    Dim lnPorcentaje_Impuesto As Decimal = CDec(loFila("Porcentaje_Impuesto"))
                    Dim lcNumero_Expediente As String = CStr(loFila("Numero_Expediente")).Trim()

                    loSalida.Append(lcRif).Append(vbTab)
                    loSalida.Append(lcPeriodo).Append(vbTab)
                    loSalida.Append(Strings.Format(ldFecha_Retencion, "yyyy-MM-dd")).Append(vbTab)
                    loSalida.Append(lcTipo_Operacion).Append(vbTab)
                    loSalida.Append(lcTipo_Documento).Append(vbTab)
                    loSalida.Append(lcRif_Proveedor).Append(vbTab)
                    loSalida.Append(lcDocumento_Origen).Append(vbTab)
                    loSalida.Append(lcControl_Origen).Append(vbTab)
                    loSalida.Append(goServicios.mObtenerFormatoCadenaCSV(lnMonto_Neto, goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 2)).Append(vbTab)
                    loSalida.Append(goServicios.mObtenerFormatoCadenaCSV(lnBase_Imponible, goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 2)).Append(vbTab)
                    loSalida.Append(goServicios.mObtenerFormatoCadenaCSV(lnMonto_Retencion, goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 2)).Append(vbTab)
                    loSalida.Append(lcNumero_Factura).Append(vbTab)
                    loSalida.Append(lcNumero_Comprobante).Append(vbTab)
                    loSalida.Append(goServicios.mObtenerFormatoCadenaCSV(lnMonto_Exento, goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 2)).Append(vbTab)
                    loSalida.Append(goServicios.mObtenerFormatoCadenaCSV(lnPorcentaje_Impuesto, goServicios.enuOpcionesRedondeo.KN_RedondeoPuntoMedio, 2)).Append(vbTab)
                    loSalida.Append(lcNumero_Expediente).Append(vbTab)
                    loSalida.AppendLine()

                Next loFila

            End If



            '-------------------------------------------------------------------------------------------------------
            ' Envia la salida a pantalla en un archivo descargable.
            '-------------------------------------------------------------------------------------------------------
            Me.Response.Clear()
            Me.Response.ContentEncoding = System.Text.Encoding.UTF8
            Me.Response.AppendHeader("content-disposition", "attachment; filename=RelacionRetencionesIVA" & lcPeriodo & ".txt")
            Me.Response.ContentType = "text/plain"
            Me.Response.Write(loSalida.ToString())
            'Me.Response.Write(Strings.Space(20))	'A veces no todo el texto es enviado a pantalla, entonces se 
            Me.Response.End()                       'mandan algunos espacios en blanco adicionales para "rellenar".

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

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 01/08/09: Codigo inicial (creado a partir de rCRetencion_IVAProveedores)				'
'-------------------------------------------------------------------------------------------'
' JJD: 08/06/11: Ajustes en la generacion de monto base y porcentaje.
'-------------------------------------------------------------------------------------------'
' MAT: 10/06/11: Generación del archivo del comprobante con cero retención
'-------------------------------------------------------------------------------------------'
' RJG: 16/04/13: Corrección de bug: algunos documentos aparecían duplicados.                
'-------------------------------------------------------------------------------------------'
' RJG: 04/06/13: Se agregó la eliminación de todos los carecteres no válidos del RIF.       '
'-------------------------------------------------------------------------------------------'
