'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRDetallada_ISLRRProveedoresXML"
'-------------------------------------------------------------------------------------------'
Partial Class rRDetallada_ISLRRProveedoresXML
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
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()




            loConsulta.AppendLine("SELECT		Cuentas_Pagar.Tip_Ori					                    AS Tipo_Origen,							")
            loConsulta.AppendLine("				Renglones_Pagos.Doc_Ori					                    AS Documento_Origen,					")
            loConsulta.AppendLine("				Renglones_Pagos.Factura		                                AS Factura_Origen,")
            loConsulta.AppendLine("				Renglones_Pagos.Control					                    AS Control_Origen,						")
            loConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini					                    AS Fecha_Retencion,						")
            loConsulta.AppendLine("				COALESCE(Detalles_Pagos.Tip_Ope, '-')	                    AS Tipo_Pago,							")
            loConsulta.AppendLine("				CASE WHEN COALESCE(Detalles_Pagos.Tip_Ope, '-')='Efectivo'						")
            loConsulta.AppendLine("					THEN 'Efectivo'																")
            loConsulta.AppendLine("					ELSE COALESCE(Detalles_Pagos.Num_Doc,'-')										")
            loConsulta.AppendLine("				END										                    AS Numero_Pago,							")
			loConsulta.AppendLine("				Renglones_Pagos.Mon_Abo					                    AS Monto_Abonado,						")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas			                    AS Base_Retencion,						")
            loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret			                    AS Porcentaje_Retenido,					")
            loConsulta.AppendLine("				Retenciones.Cod_Ret						                    AS Codigo_Concepto,						")
            loConsulta.AppendLine("				Retenciones.Nom_Ret						                    AS Concepto,							")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret			                    AS Monto_Retenido,						")
            loConsulta.AppendLine("				Cuentas_Pagar.Cod_Pro					                    AS Cod_Pro,								")
            loConsulta.AppendLine("				SUBSTRING(Proveedores.Nom_Pro, 0,30)	                    AS Nom_Pro,								")
            loConsulta.AppendLine("				Proveedores.Rif							                    AS Rif,									")
            loConsulta.AppendLine("				SUBSTRING(Proveedores.Dir_Fis, 0,25)	                    AS Direccion							")
            loConsulta.AppendLine("FROM			Cuentas_Pagar")
            loConsulta.AppendLine("		JOIN	Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loConsulta.AppendLine("		JOIN	Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loConsulta.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine("		LEFT JOIN Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento")
            loConsulta.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("		JOIN	Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")

            loConsulta.AppendLine("           AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("         		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("           AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("         		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("           AND Pagos.Cod_Mon BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("         		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("           AND Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("         		AND " & lcParametro3Hasta)

            loConsulta.AppendLine("UNION ALL		")
            
            loConsulta.AppendLine("SELECT	Retenciones_Documentos.Tip_Ori				AS Tipo_Origen,")
            loConsulta.AppendLine("			Ordenes_Pagos.Documento						AS Documento_Origen,")
            loConsulta.AppendLine("			Ordenes_Pagos.Documento		                AS Factura_Origen,")
            loConsulta.AppendLine("			Ordenes_Pagos.Control						AS Control_Origen,")
            loConsulta.AppendLine("			Ordenes_Pagos.Fec_Ini						AS Fecha_Retencion,")
            loConsulta.AppendLine("			Detalles_OPagos.Tip_Ope						AS Tipo_Pago,")
            loConsulta.AppendLine("			CASE WHEN Detalles_OPagos.Tip_Ope='Efectivo'")
            loConsulta.AppendLine("				THEN 'Efectivo'				")
            loConsulta.AppendLine("				ELSE Detalles_OPagos.Num_Doc")
            loConsulta.AppendLine("			END											AS Numero_Pago,	")
            loConsulta.AppendLine("			Ordenes_Pagos.Mon_Net						AS Monto_Abonado,")
            loConsulta.AppendLine("			Retenciones_Documentos.Mon_Bas				AS Base_Retencion,")
            loConsulta.AppendLine("			Retenciones_Documentos.Por_Ret				AS Porcentaje_Retenido,")
            loConsulta.AppendLine("			Retenciones.Cod_Ret							AS Codigo_Concepto,")
            loConsulta.AppendLine("			Retenciones.Nom_Ret							AS Concepto,")
            loConsulta.AppendLine("			Retenciones_Documentos.Mon_Ret				AS Monto_Retenido,")
            loConsulta.AppendLine("			Ordenes_Pagos.Cod_Pro						AS Cod_Pro,")
            loConsulta.AppendLine("			SUBSTRING(Proveedores.Nom_Pro, 0,30)		AS Nom_Pro,") 
            loConsulta.AppendLine("			Proveedores.Rif								AS Rif,") 
            loConsulta.AppendLine("			SUBSTRING(Proveedores.Dir_Fis, 0,25)		AS Direccion")
            loConsulta.AppendLine("FROM		Retenciones_Documentos")
            loConsulta.AppendLine("	JOIN	Ordenes_Pagos ON Ordenes_Pagos.Documento = Retenciones_Documentos.documento")
            loConsulta.AppendLine("	JOIN	Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            loConsulta.AppendLine("	JOIN	Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("	JOIN	Detalles_OPagos ON Detalles_OPagos.Documento = Ordenes_Pagos.Documento")
            loConsulta.AppendLine("WHERE		Ordenes_Pagos.Status = 'Confirmado'")
            loConsulta.AppendLine("			AND	Retenciones_Documentos.Tip_Ori = 'Ordenes_Pagos'")

            loConsulta.AppendLine("           AND Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("         		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("           AND Ordenes_Pagos.Cod_Pro BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("         		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("           AND Ordenes_Pagos.Cod_Mon BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("         		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("           AND Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("         		AND " & lcParametro3Hasta)
            

            loConsulta.AppendLine("UNION ALL		")
            
            loConsulta.AppendLine("SELECT		Cuentas_Pagar.Tip_Ori					AS Tipo_Origen,")
            loConsulta.AppendLine("				Cuentas_Pagar.Doc_Ori					AS Documento_Origen,")
            loConsulta.AppendLine("				Documentos.Factura		                AS Factura_Origen,")
            loConsulta.AppendLine("				Documentos.Control						AS Control_Origen,")
			loConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini					AS Fecha_Retencion,")
			loConsulta.AppendLine("				'-'										AS Tipo_Pago,")
			loConsulta.AppendLine("				'-'										AS Numero_Pago,")
			loConsulta.AppendLine("				Documentos.Mon_Net						AS Monto_Abonado,")
			loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas			AS Base_Retencion,")
			loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret			AS Porcentaje_Retenido,")
			loConsulta.AppendLine("				Retenciones.Cod_Ret						AS Codigo_Concepto,") 
			loConsulta.AppendLine("				Retenciones.Nom_Ret						AS Concepto,") 
			loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret			AS Monto_Retenido,")
			loConsulta.AppendLine("				Cuentas_Pagar.Cod_Pro					AS Cod_Pro,")
			loConsulta.AppendLine("				SUBSTRING(Proveedores.Nom_Pro, 0,30)	AS Nom_Pro,")
			loConsulta.AppendLine("				Proveedores.Rif							AS Rif,")
			loConsulta.AppendLine("				SUBSTRING(Proveedores.Dir_Fis, 0,25)	AS Direccion")
			loConsulta.AppendLine("FROM			Cuentas_Pagar")
			loConsulta.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("		JOIN	Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'") 
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'") 

            loConsulta.AppendLine("       	AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("       		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("       	AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("       		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("       	AND Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("       		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("       	AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("       		AND " & lcParametro3Hasta)

            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
			
        '-------------------------------------------------------------------------------------------------------
        ' Genera el XML
        '-------------------------------------------------------------------------------------------------------
			Dim loSalida As New StringBuilder()
			
			Dim lcRifEmpresa	As String = Strings.Trim(goEmpresa.pcRifEmpresa).Replace("-", "").Replace(".", "").Replace(" ", "")
			Dim ldFecha			As Date = CDate(cusAplicacion.goReportes.paParametrosIniciales(0))
			Dim lcPeriodo		As String = Strings.Format(ldFecha, "yyyy") & Strings.Format(ldFecha, "MM")
			
			loSalida.AppendLine("<?xml version=""1.0"" encoding=""utf-8"" ?>")
			loSalida.Append("<RelacionRetencionesISLR RifAgente=""").Append(lcRifEmpresa).Append(""" Periodo=""").Append(lcPeriodo).AppendLine(""" >")	
			
			Const KC_10Ceros	As String = "0000000000"
			Const KC_08Ceros	As String = "00000000"
			Const KC_03Ceros	As String = "000"
			Const KC_03Espacios	As String = "   "
			Const KC_13Espacios As String = "             "
			Const KC_06Espacios As String = "      "
				  
			For Each loRenglon As DataRow IN laDatosReporte.Tables(0).Rows
				
			'Numero de RIF con formato: "A00000000" (sin guiones, espacios o puntos)
				Dim lcRifRetenido		As String = Strings.Trim(loRenglon("Rif")).Replace("-", "").Replace(".", "").Replace(" ", "")
				
			'Numero de Factura con formato: "0000000000" (10 caracteres, relleno con ceros a la izquierda)
				Dim lcNumeroFactura As String = Strings.Trim(loRenglon("Factura_Origen"))
				lcNumeroFactura = IIf(lcNumeroFactura = "0", "0         ", Strings.Right(KC_10Ceros & lcNumeroFactura, 10))
				
			'Numero de Control con formato: "00000000" (8 caracteres, relleno con ceros a la izquierda)
				Dim lcNumeroControl As String = Regex.Replace(Strings.Trim(loRenglon("Control_Origen")), "[^0-9]", "") 
				lcNumeroControl = IIf(String.IsNullOrEmpty(lcNumeroControl), "NA      ", Strings.Right(KC_08Ceros & lcNumeroControl, 8))
				
			'Fecha de la Operación con formato: dd/mm/aaaa (10 caracteres)
				Dim ldFechaOperacion As Date = CDate(loRenglon("Fecha_Retencion"))
				Dim lcFechaOperacion As String = ldFechaOperacion.ToString("dd/MM/yyyy")
				
			'Código de Concepto con formato: "AAA" (3 caracteres, relleno con espacios a la derecha)
				Dim lcCodigoConcepto As String = Strings.Trim(loRenglon("Codigo_Concepto"))
				lcCodigoConcepto = Strings.Right(KC_03Ceros & lcCodigoConcepto, 3)
				
			'Monto base de la retención: redondeado a dos decimales, con 13 caracteres, relleno con espacios a la izquierda
				Dim lnMontoOperacion As Decimal = goServicios.mRedondearValor(CDec(loRenglon("Base_Retencion")), 2)
				Dim lcMontoOperacion As String = Strings.Right(KC_13Espacios & Strings.Format(lnMontoOperacion, "0.00"), 13)
				
			'Porcentaje de retención: redondeado a dos decimales, con 6 caracteres, relleno con espacios a la izquierda
				Dim lnPorcentajeRetenido As Decimal = goServicios.mRedondearValor(CDec(loRenglon("Monto_Retenido"))*100/CDec(loRenglon("Base_Retencion")), 2)
				Dim lcPorcentajeRetenido As String = Strings.Right(KC_06Espacios & Strings.Format(lnPorcentajeRetenido, "0.00"), 6)
				
				loSalida.Append("<DetalleRetencion>")
				loSalida.Append(	"<RifRetenido>").Append(lcRifRetenido).Append("</RifRetenido>")
				loSalida.Append(	"<NumeroFactura>").Append(lcNumeroFactura).Append("</NumeroFactura>")
				loSalida.Append(	"<NumeroControl>").Append(lcNumeroControl).Append("</NumeroControl>")
				loSalida.Append(	"<FechaOperacion>").Append(lcFechaOperacion).Append("</FechaOperacion>")
				loSalida.Append(	"<CodigoConcepto>").Append(lcCodigoConcepto).Append("</CodigoConcepto>")
				loSalida.Append(	"<MontoOperacion>").Append(lcMontoOperacion).Append("</MontoOperacion>")
				loSalida.Append(	"<PorcentajeRetencion>").Append(lcPorcentajeRetenido).Append("</PorcentajeRetencion>")
				loSalida.Append("</DetalleRetencion>")
				loSalida.AppendLine()				
				
			Next loRenglon
			
			loSalida.AppendLine("</RelacionRetencionesISLR>")
			


			
			
			Me.Response.Clear()
			Me.Response.ContentEncoding = System.Text.Encoding.UTF8
			Me.Response.AppendHeader("content-disposition", "attachment; filename=RelacionRetencionesISLR" & lcPeriodo & ".xml")
			Me.Response.ContentType = "application/xml"
			Me.Response.Write(loSalida.ToString())
			'Me.Response.Write(Strings.Space(20))
            Me.Response.Flush()
			Me.Response.End()
			

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
' RJG: 31/05/11: Codigo inicial (partir del reporte "rRDetallada_ISLRRProveedores").		'
'-------------------------------------------------------------------------------------------'
' RJG: 04/07/11: Eliminados caracteres que no sean números en el número de control. Relleno	'
'				 del código de concepto con ceros a la izquierda en lugar de espacios a la	'
'				 derecha.																	'
'-------------------------------------------------------------------------------------------'
' RJG: 06/07/11: Agregado LEFT al JOIN de Detalles de Pagos (no siempre hay detalle).		'
'-------------------------------------------------------------------------------------------'
' RJG: 05/09/15: Se cambió el número de documento (interno) por el número de factura (del   '
'                proveedor).                                                                '
'-------------------------------------------------------------------------------------------'
