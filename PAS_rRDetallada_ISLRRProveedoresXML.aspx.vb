'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "PAS_rRDetallada_ISLRRProveedoresXML"
'-------------------------------------------------------------------------------------------'
Partial Class PAS_rRDetallada_ISLRRProveedoresXML
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("DECLARE @lnCero AS DECIMAL(28, 10) = CAST(0 AS DECIMAL(28, 10));")
            loConsulta.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)

            loConsulta.AppendLine("SELECT   Renglones_Pagos.Factura										AS Factura_Origen,")
            loConsulta.AppendLine("         Renglones_Pagos.Control					                    AS Control_Origen,	")
            loConsulta.AppendLine("         Cuentas_Pagar.Fec_Ini					                    AS Fecha_Retencion,")
            'loConsulta.AppendLine("         Retenciones_Documentos.Mon_Bas			                    AS Base_Retencion,")
            loConsulta.AppendLine("         CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini >= '20180801'")
            loConsulta.AppendLine("              THEN (Retenciones_Documentos.Mon_Bas/100000)")
            loConsulta.AppendLine("              ELSE Retenciones_Documentos.Mon_Bas")
            loConsulta.AppendLine("         END                                                         AS Base_Retencion,")
            loConsulta.AppendLine("         Retenciones_Documentos.Por_Ret			                    AS Porcentaje_Retenido,")
            loConsulta.AppendLine("         RIGHT(CONCAT('00',RTRIM(Retenciones.Cod_Ret)),3)            AS Codigo_Concepto,	")
            loConsulta.AppendLine("         Retenciones_Documentos.Mon_Ret			                    AS Monto_Retenido,")
            loConsulta.AppendLine("         Proveedores.Rif							                    AS Rif")
            loConsulta.AppendLine("FROM Cuentas_Pagar")
            loConsulta.AppendLine(" JOIN Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine(" JOIN Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loConsulta.AppendLine("     AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loConsulta.AppendLine(" JOIN Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loConsulta.AppendLine("     AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine(" LEFT JOIN Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento")
            loConsulta.AppendLine(" JOIN Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine(" JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loConsulta.AppendLine(" AND	Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine(" AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")
            loConsulta.AppendLine(" AND Cuentas_Pagar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   Documentos.Factura											AS Factura_Origen,")
            loConsulta.AppendLine("         Documentos.Control											AS Control_Origen,")
            loConsulta.AppendLine("         Cuentas_Pagar.Fec_Ini										AS Fecha_Retencion,")
            'loConsulta.AppendLine("         Retenciones_Documentos.Mon_Bas								AS Base_Retencion,")
            loConsulta.AppendLine("         CASE WHEN Cuentas_Pagar.Fec_Ini < '20180820' AND Cuentas_Pagar.Fec_Ini >= '20180801'")
            loConsulta.AppendLine("              THEN (Retenciones_Documentos.Mon_Bas/100000)")
            loConsulta.AppendLine("              ELSE Retenciones_Documentos.Mon_Bas")
            loConsulta.AppendLine("         END                                                         AS Base_Retencion,")
            loConsulta.AppendLine("         Retenciones_Documentos.Por_Ret								AS Porcentaje_Retenido,")
            loConsulta.AppendLine("         RIGHT(CONCAT('00',RTRIM(Retenciones.Cod_Ret)),3)			AS Codigo_Concepto,")
            loConsulta.AppendLine("         Retenciones_Documentos.Mon_Ret								AS Monto_Retenido,")
            loConsulta.AppendLine("         Proveedores.Rif												AS Rif")
            loConsulta.AppendLine("FROM Cuentas_Pagar")
            loConsulta.AppendLine(" JOIN Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("     AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loConsulta.AppendLine(" JOIN Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loConsulta.AppendLine("     AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine(" JOIN Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine(" JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loConsulta.AppendLine(" AND Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine(" AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            loConsulta.AppendLine(" AND Cuentas_Pagar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	CASE WHEN Trabajadores.Cod_Con = '02'")
            loConsulta.AppendLine("			     THEN REPLACE(CONVERT(VARCHAR(10),Recibos.Fecha, 103),'/','')")
            loConsulta.AppendLine("			     ELSE CASE WHEN DAY(Recibos.Fecha) <= 15")
            loConsulta.AppendLine("                        THEN CONCAT('15',REPLACE(RIGHT(CONVERT(VARCHAR(10),Recibos.Fecha, 103), 7),'/',''))")
            loConsulta.AppendLine("						   ELSE CONCAT(CAST(DAY(DATEADD(MS,- 3,DATEADD(MM,0,DATEADD(M,DATEDIFF(MM,0,Recibos.Fecha)+1,0)))) AS VARCHAR(2)),REPLACE(RIGHT(CONVERT(VARCHAR(10),Recibos.Fecha, 103), 7),'/',''))")
            loConsulta.AppendLine("                   END	")
            loConsulta.AppendLine("		    END																	AS Factura_Origen,")
            loConsulta.AppendLine("		    CASE WHEN Trabajadores.Cod_Con = '02'")
            loConsulta.AppendLine("			     THEN REPLACE(CONVERT(VARCHAR(10),Recibos.Fecha, 103),'/','')")
            loConsulta.AppendLine("			     ELSE CASE WHEN DAY(Recibos.Fecha) <= 15")
            loConsulta.AppendLine("					       THEN CONCAT('15',REPLACE(RIGHT(CONVERT(VARCHAR(10),Recibos.Fecha, 103), 7),'/',''))")
            loConsulta.AppendLine("					       ELSE CONCAT(CAST(DAY(DATEADD(MS,- 3,DATEADD(MM,0,DATEADD(M,DATEDIFF(MM,0,Recibos.Fecha)+1,0)))) AS VARCHAR(2)),REPLACE(RIGHT(CONVERT(VARCHAR(10),Recibos.Fecha, 103), 7),'/',''))")
            loConsulta.AppendLine("				      END	")
            loConsulta.AppendLine("		    END															        AS Control_Origen,")
            loConsulta.AppendLine("	        Recibos.Fecha												        AS Fecha_Retencion,")
            'loConsulta.AppendLine("         COALESCE(ROUND(Renglones_Recibos.Mon_Net * 100 / CAST (SUBSTRING(Renglones_Recibos.val_car,0, LEN(Renglones_Recibos.val_car)) AS DECIMAL (28,2)),2)")
            'loConsulta.AppendLine("			,(SELECT SUM(Mon_Net)")
            'loConsulta.AppendLine("			    FROM Renglones_Recibos")
            'loConsulta.AppendLine("			    WHERE Documento = Recibos.Documento")
            'loConsulta.AppendLine("				    AND Tipo = 'Asignacion'),@lnCero)					        AS Base_Retencion,")
            loConsulta.AppendLine("			COALESCE((SELECT CASE WHEN Recibos.Fecha < '20180820' AND Recibos.Fecha >= '20180801'")
            loConsulta.AppendLine("                               THEN ((ROUND(Renglones_Recibos.Mon_Net * 100 / CAST (SUBSTRING(Renglones_Recibos.val_car,0, LEN(Renglones_Recibos.val_car)) AS DECIMAL (28,2)),2))/100000)")
            loConsulta.AppendLine("                               ELSE ROUND(Renglones_Recibos.Mon_Net * 100 / CAST (SUBSTRING(Renglones_Recibos.val_car,0, LEN(Renglones_Recibos.val_car)) AS DECIMAL (28,2)),2)")
            loConsulta.AppendLine("                         END")
            loConsulta.AppendLine("				      FROM Renglones_Recibos")
            loConsulta.AppendLine("				      WHERE Documento = Recibos.Documento")
            loConsulta.AppendLine("					    AND Cod_Con IN ('R005', 'R405')")
            loConsulta.AppendLine("			),(SELECT SUM(CASE WHEN Renglones_Recibos.Fec_Ini < '20180820' AND Renglones_Recibos.Fec_Ini >= '20180801'")
            loConsulta.AppendLine("                               THEN (Renglones_Recibos.Mon_Net/100000)")
            loConsulta.AppendLine("                               ELSE Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("                         END)")
            loConsulta.AppendLine("				FROM Renglones_Recibos")
            loConsulta.AppendLine("				WHERE Documento = Recibos.Documento")
            loConsulta.AppendLine("					AND Tipo = 'Asignacion'),@lnCero)						AS Base_Retencion,")
            loConsulta.AppendLine("         COALESCE(CAST (SUBSTRING(Renglones_Recibos.val_car,0, LEN(Renglones_Recibos.val_car)) AS DECIMAL (28,2)), @lnCero)  AS Porcentaje_Retenido,")
            loConsulta.AppendLine("	        '001'														        AS Codigo_Concepto,")
            loConsulta.AppendLine("         COALESCE(Renglones_Recibos.Mon_Net, @lnCero)				        AS Monto_Retenido,")
            loConsulta.AppendLine("	        Trabajadores.Rif											        AS Rif ")
            loConsulta.AppendLine("FROM Recibos")
            loConsulta.AppendLine(" JOIN Trabajadores ON Recibos.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine(" LEFT JOIN Renglones_Recibos ON Recibos.Documento = Renglones_Recibos.Documento")
            loConsulta.AppendLine("     AND Renglones_Recibos.Cod_Con IN ('R005', 'R405')")
            loConsulta.AppendLine("WHERE ((Recibos.Fecha >= @ldFecha_Desde")
            loConsulta.AppendLine(" AND Recibos.Fecha < DATEADD(dd, DATEDIFF(dd, 0, @ldFecha_Hasta) + 1, 0))")
            loConsulta.AppendLine(" OR Renglones_Recibos.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta)")
            loConsulta.AppendLine(" AND Trabajadores.Rif <> ''")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("ORDER BY Fecha_Retencion")

            'Me.mEscribirConsulta(loConsulta.ToString())


            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Genera el XML
            '-------------------------------------------------------------------------------------------------------
            Dim loSalida As New StringBuilder()

            Dim lcRifEmpresa As String = Strings.Trim(goEmpresa.pcRifEmpresa).Replace("-", "").Replace(".", "").Replace(" ", "")
            Dim ldFecha As Date = CDate(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcPeriodo As String = Strings.Format(ldFecha, "yyyy") & Strings.Format(ldFecha, "MM")

            loSalida.AppendLine("<?xml version=""1.0"" encoding=""utf-8"" ?>")
            loSalida.Append("<RelacionRetencionesISLR RifAgente=""").Append(lcRifEmpresa).Append(""" Periodo=""").Append(lcPeriodo).AppendLine(""" >")

            Const KC_10Ceros As String = "0000000000"
            Const KC_08Ceros As String = "00000000"
            Const KC_03Ceros As String = "000"
            Const KC_03Espacios As String = "   "
            Const KC_13Espacios As String = "             "
            Const KC_06Espacios As String = "      "

            For Each loRenglon As DataRow In laDatosReporte.Tables(0).Rows

                'Numero de RIF con formato: "A00000000" (sin guiones, espacios o puntos)
                Dim lcRifRetenido As String = Strings.Trim(loRenglon("Rif")).Replace("-", "").Replace(".", "").Replace(" ", "")

                'Numero de Factura con formato: "0000000000" (10 caracteres, relleno con ceros a la izquierda)
                Dim lcNumeroFactura As String = Strings.Trim(loRenglon("Factura_Origen").Replace("-", ""))
                If lcNumeroFactura <> "NA" Then
                    lcNumeroFactura = IIf(lcNumeroFactura = "0", "0         ", Strings.Right(KC_10Ceros & lcNumeroFactura, 10))
                End If

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
                Dim lnPorcentajeRetenido As Decimal = goServicios.mRedondearValor(CDec(loRenglon("Porcentaje_Retenido")), 2)
                Dim lcPorcentajeRetenido As String = Strings.Right(KC_06Espacios & Strings.Format(lnPorcentajeRetenido, "0.00"), 6)

                loSalida.Append("<DetalleRetencion>")
                loSalida.Append("<RifRetenido>").Append(lcRifRetenido).Append("</RifRetenido>")
                loSalida.Append("<NumeroFactura>").Append(lcNumeroFactura).Append("</NumeroFactura>")
                loSalida.Append("<NumeroControl>").Append(lcNumeroControl).Append("</NumeroControl>")
                loSalida.Append("<FechaOperacion>").Append(lcFechaOperacion).Append("</FechaOperacion>")
                loSalida.Append("<CodigoConcepto>").Append(lcCodigoConcepto).Append("</CodigoConcepto>")
                loSalida.Append("<MontoOperacion>").Append(lcMontoOperacion).Append("</MontoOperacion>")
                loSalida.Append("<PorcentajeRetencion>").Append(lcPorcentajeRetenido).Append("</PorcentajeRetencion>")
                loSalida.Append("</DetalleRetencion>")
                loSalida.AppendLine()

            Next loRenglon

            loSalida.AppendLine("</RelacionRetencionesISLR>")

            Me.Response.Clear()
            Me.Response.ContentEncoding = System.Text.Encoding.UTF8
            Me.Response.AppendHeader("content-disposition", "attachment; filename=RelacionRetencionesISLR_PAS_" & lcPeriodo & ".xml")
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
