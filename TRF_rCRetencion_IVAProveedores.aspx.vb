'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rCRetencion_IVAProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rCRetencion_IVAProveedores
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

            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("               Cuentas_Pagar.Documento             AS Documento,")
            loComandoSeleccionar.AppendLine("               Cuentas_Pagar.Doc_Ori               AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Num_Com		AS Numero_Comprobante,			")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Documento				AS Numero_DocumentoRet,			")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Control				AS Numero_ControlRet,			")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Control					AS Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("				Documentos.Factura					AS Numero_FacturaDoc,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Mon_Net				AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Exe					AS Monto_ExentoDoc,				")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Bas1					AS Monto_BaseDoc,				")
            loComandoSeleccionar.AppendLine("				Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,		")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Imp1					AS Monto_ImpuestoDoc,			")
            loComandoSeleccionar.AppendLine("				Documentos.Fec_Ini					AS Fecha_Recepcion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Proveedores.Dir_Fis					AS Direccion,")
            loComandoSeleccionar.AppendLine("				YEAR(Cuentas_Pagar.Fec_Ini)			AS Anio,")
            loComandoSeleccionar.AppendLine("				MONTH(Cuentas_Pagar.Fec_Ini)		AS Mes")
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
            loComandoSeleccionar.AppendLine("		LEFT JOIN	Cuentas_Pagar AS Documentos										")
            loComandoSeleccionar.AppendLine("			ON	Documentos.Documento = Renglones_Pagos.Doc_Ori					")
            loComandoSeleccionar.AppendLine("			AND	Documentos.Cod_Tip = Renglones_Pagos.Cod_Tip 					")
            loComandoSeleccionar.AppendLine("		JOIN	Proveedores")
            loComandoSeleccionar.AppendLine("			ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("		LEFT JOIN Retenciones")
            loComandoSeleccionar.AppendLine("			ON	Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")

            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Documentos.Factura BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		")

            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("               Cuentas_Pagar.Documento             AS Documento,")
            loComandoSeleccionar.AppendLine("               Cuentas_Pagar.Doc_Ori               AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Num_Com		AS Numero_Comprobante,				")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Documento				AS Numero_DocumentoRet,				")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Control				AS Numero_ControlRet,				")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Control					AS Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("               Documentos.Factura                  AS Numero_FacturaDoc,")
            loComandoSeleccionar.AppendLine("               CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("                    THEN Documentos.Mon_Net * (-1)")
            loComandoSeleccionar.AppendLine("                    ELSE Documentos.Mon_Net")
            loComandoSeleccionar.AppendLine("               END                                 AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("               CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("                    THEN Documentos.Mon_Exe * (-1)")
            loComandoSeleccionar.AppendLine("                    ELSE Documentos.Mon_Exe")
            loComandoSeleccionar.AppendLine("               END                                 AS Monto_ExentoDoc,")
            loComandoSeleccionar.AppendLine("               CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("                    THEN Documentos.Mon_Bas1 * (-1)")
            loComandoSeleccionar.AppendLine("                    ELSE Documentos.Mon_Bas1")
            loComandoSeleccionar.AppendLine("               END                                 AS Monto_BaseDoc,")
            loComandoSeleccionar.AppendLine("				Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,			")
            loComandoSeleccionar.AppendLine("               CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("                    THEN Documentos.Mon_Imp1 * (-1)")
            loComandoSeleccionar.AppendLine("                    ELSE Documentos.Mon_Imp1")
            loComandoSeleccionar.AppendLine("               END                                 AS Monto_ImpuestoDoc,")
            loComandoSeleccionar.AppendLine("				Documentos.Fec_Ini					AS Fecha_Recepcion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loComandoSeleccionar.AppendLine("               CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("                    THEN Retenciones_Documentos.Mon_Ret * (-1)")
            loComandoSeleccionar.AppendLine("                    ELSE Retenciones_Documentos.Mon_Ret")
            loComandoSeleccionar.AppendLine("               END                                 AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Proveedores.Dir_Fis					AS Direccion,")
            loComandoSeleccionar.AppendLine("				YEAR(Cuentas_Pagar.Fec_Ini)			AS Anio,")
            loComandoSeleccionar.AppendLine("				MONTH(Cuentas_Pagar.Fec_Ini)		AS Mes")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ")
            loComandoSeleccionar.AppendLine("			ON	Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
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
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Documentos.Factura BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro2Hasta)

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
            loComandoSeleccionar.AppendLine("                                    ELSE #tabRetenciones.Rif ")
            loComandoSeleccionar.AppendLine("                                   END),")
            loComandoSeleccionar.AppendLine("      #tabRetenciones.Nom_Pro = (SELECT CASE WHEN (ISNULL((SELECT Cod_Ter")
            loComandoSeleccionar.AppendLine("                                                      FROM Retenciones_Renglones")
            loComandoSeleccionar.AppendLine("                                                      WHERE Doc_Des = #tabRetenciones.Documento   ")
            loComandoSeleccionar.AppendLine("                                                      AND Doc_Ori = #tabRetenciones.Doc_Ori),'') <>'')  ")
            loComandoSeleccionar.AppendLine("                                    THEN (SELECT Nom_Pro")
            loComandoSeleccionar.AppendLine("                                          FROM Proveedores")
            loComandoSeleccionar.AppendLine("                                          WHERE Cod_Pro = (SELECT Cod_Ter")
            loComandoSeleccionar.AppendLine("                                                                  FROM Retenciones_Renglones")
            loComandoSeleccionar.AppendLine("                                                                  WHERE Doc_Des = #tabRetenciones.Documento")
            loComandoSeleccionar.AppendLine("                                                                  AND Doc_Ori = #tabRetenciones.Doc_Ori))")
            loComandoSeleccionar.AppendLine("                                    ELSE #tabRetenciones.Nom_Pro ")
            loComandoSeleccionar.AppendLine("                                   END),")
            loComandoSeleccionar.AppendLine("      #tabRetenciones.Nit = (SELECT CASE WHEN (ISNULL((SELECT Cod_Ter")
            loComandoSeleccionar.AppendLine("                                                      FROM Retenciones_Renglones")
            loComandoSeleccionar.AppendLine("                                                      WHERE Doc_Des = #tabRetenciones.Documento   ")
            loComandoSeleccionar.AppendLine("                                                      AND Doc_Ori = #tabRetenciones.Doc_Ori),'') <>'')  ")
            loComandoSeleccionar.AppendLine("                                    THEN (SELECT Nit")
            loComandoSeleccionar.AppendLine("                                          FROM Proveedores")
            loComandoSeleccionar.AppendLine("                                          WHERE Cod_Pro = (SELECT Cod_Ter")
            loComandoSeleccionar.AppendLine("                                                                  FROM Retenciones_Renglones")
            loComandoSeleccionar.AppendLine("                                                                  WHERE Doc_Des = #tabRetenciones.Documento")
            loComandoSeleccionar.AppendLine("                                                                  AND Doc_Ori = #tabRetenciones.Doc_Ori))")
            loComandoSeleccionar.AppendLine("                                    ELSE #tabRetenciones.Nit ")
            loComandoSeleccionar.AppendLine("                                   END),")
            loComandoSeleccionar.AppendLine("      #tabRetenciones.Direccion = (SELECT CASE WHEN (ISNULL((SELECT Cod_Ter")
            loComandoSeleccionar.AppendLine("                                                      FROM Retenciones_Renglones")
            loComandoSeleccionar.AppendLine("                                                      WHERE Doc_Des = #tabRetenciones.Documento   ")
            loComandoSeleccionar.AppendLine("                                                      AND Doc_Ori = #tabRetenciones.Doc_Ori),'') <>'')  ")
            loComandoSeleccionar.AppendLine("                                    THEN (SELECT Dir_Fis")
            loComandoSeleccionar.AppendLine("                                          FROM Proveedores")
            loComandoSeleccionar.AppendLine("                                          WHERE Cod_Pro = (SELECT Cod_Ter")
            loComandoSeleccionar.AppendLine("                                                                  FROM Retenciones_Renglones")
            loComandoSeleccionar.AppendLine("                                                                  WHERE Doc_Des = #tabRetenciones.Documento")
            loComandoSeleccionar.AppendLine("                                                                  AND Doc_Ori = #tabRetenciones.Doc_Ori))")
            loComandoSeleccionar.AppendLine("                                    ELSE #tabRetenciones.Direccion ")
            loComandoSeleccionar.AppendLine("                                   END)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT #tabRetenciones.Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Numero_Comprobante,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Numero_FacturaDoc,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Monto_Documento,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Monto_ExentoDoc,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Monto_BaseDoc,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Porcentaje_ImpuestoDoc,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Monto_ImpuestoDoc,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Fecha_Recepcion,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Base_Retencion,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Monto_Retenido,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Nom_Pro,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Rif,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Nit,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Direccion,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Anio,")
            loComandoSeleccionar.AppendLine("        #tabRetenciones.Mes")
            loComandoSeleccionar.AppendLine("FROM #tabRetenciones")
            loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento & ", Fecha_Retencion ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tabRetenciones")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº 0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rCRetencion_IVAProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTRF_rCRetencion_IVAProveedores.ReportSource = loObjetoReporte


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
' Fin del codigo
'-------------------------------------------------------------------------------------------'
' RJG: 09/04/15: Codigo inicial, a partir de rCRetencion_IVAProveedores.                    '
'-------------------------------------------------------------------------------------------'
