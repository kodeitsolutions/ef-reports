'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rResumenRetencion_IVAProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rResumenRetencion_IVAProveedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Cuentas_Pagar.Documento				AS Documento,")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Doc_Ori				AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Fec_Ini				AS Fecha_Doc,")
            loComandoSeleccionar.AppendLine("       CASE WHEN MONTH(Retenciones_Documentos.Fec_Ini) < 10")
            loComandoSeleccionar.AppendLine("           THEN CONCAT(YEAR(Retenciones_Documentos.Fec_Ini), '0', MONTH(Retenciones_Documentos.Fec_Ini), Retenciones_Documentos.Num_Com)")
            loComandoSeleccionar.AppendLine("           ELSE CONCAT(YEAR(Retenciones_Documentos.Fec_Ini), MONTH(Retenciones_Documentos.Fec_Ini), Retenciones_Documentos.Num_Com)")
            loComandoSeleccionar.AppendLine("       END									AS Numero_Comprobante,")
            loComandoSeleccionar.AppendLine("		Documentos.Control					AS Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Factura					AS Numero_FacturaDoc,")
            loComandoSeleccionar.AppendLine("		Renglones_Pagos.Mon_Net				AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("		Documentos.Mon_Exe					AS Monto_ExentoDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Mon_Bas1					AS Monto_BaseDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Fec_Ini					AS Fecha_Factura,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Rif						AS Rif, ")
            loComandoSeleccionar.AppendLine("       CAST(@ldFecha_Desde AS DATE)        AS Fecha_De,")
            loComandoSeleccionar.AppendLine("       CAST(@ldFecha_Hasta AS DATE)        AS Fecha_Hasta")
            loComandoSeleccionar.AppendLine("INTO #tabRetenciones")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("	JOIN	Pagos")
            loComandoSeleccionar.AppendLine("		ON	Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("	JOIN	Retenciones_Documentos ")
            loComandoSeleccionar.AppendLine("		ON	Retenciones_Documentos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.doc_des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.clase	= 'IMPUESTO'")
            loComandoSeleccionar.AppendLine("	JOIN	Renglones_Pagos ")
            loComandoSeleccionar.AppendLine("		ON	Renglones_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Cuentas_Pagar AS Documentos")
            loComandoSeleccionar.AppendLine("		ON	Documentos.Documento = Renglones_Pagos.Doc_Ori")
            loComandoSeleccionar.AppendLine("		AND	Documentos.Cod_Tip = Renglones_Pagos.Cod_Tip ")
            loComandoSeleccionar.AppendLine("	JOIN	Proveedores")
            loComandoSeleccionar.AppendLine("		ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Retenciones")
            loComandoSeleccionar.AppendLine("		ON	Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret ")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL	")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Cuentas_Pagar.Documento				AS Documento,")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Doc_Ori				AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Fec_Ini				AS Fecha_Doc,")
            loComandoSeleccionar.AppendLine("       CASE WHEN MONTH(Retenciones_Documentos.Fec_Ini) < 10")
            loComandoSeleccionar.AppendLine("           THEN CONCAT(YEAR(Retenciones_Documentos.Fec_Ini), '0', MONTH(Retenciones_Documentos.Fec_Ini), Retenciones_Documentos.Num_Com)")
            loComandoSeleccionar.AppendLine("           ELSE CONCAT(YEAR(Retenciones_Documentos.Fec_Ini), MONTH(Retenciones_Documentos.Fec_Ini), Retenciones_Documentos.Num_Com)")
            loComandoSeleccionar.AppendLine("       END									AS Numero_Comprobante,")
            loComandoSeleccionar.AppendLine("		Documentos.Control					AS Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Factura					AS Numero_FacturaDoc,")
            loComandoSeleccionar.AppendLine("       CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("            THEN Documentos.Mon_Net * (-1)")
            loComandoSeleccionar.AppendLine("            ELSE Documentos.Mon_Net")
            loComandoSeleccionar.AppendLine("       END                                 AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("       CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("            THEN Documentos.Mon_Exe * (-1)")
            loComandoSeleccionar.AppendLine("            ELSE Documentos.Mon_Exe")
            loComandoSeleccionar.AppendLine("       END                                 AS Monto_ExentoDoc,")
            loComandoSeleccionar.AppendLine("       CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("            THEN Documentos.Mon_Bas1 * (-1)")
            loComandoSeleccionar.AppendLine("            ELSE Documentos.Mon_Bas1")
            loComandoSeleccionar.AppendLine("       END                                 AS Monto_BaseDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Fec_Ini					AS Fecha_Factura,")
            loComandoSeleccionar.AppendLine("       CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("            THEN Retenciones_Documentos.Mon_Bas * (-1)")
            loComandoSeleccionar.AppendLine("            ELSE Retenciones_Documentos.Mon_Bas")
            loComandoSeleccionar.AppendLine("       END                                 AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("       CASE WHEN Documentos.Cod_Tip = 'N/CR'")
            loComandoSeleccionar.AppendLine("            THEN Retenciones_Documentos.Mon_Ret * (-1)")
            loComandoSeleccionar.AppendLine("            ELSE Retenciones_Documentos.Mon_Ret")
            loComandoSeleccionar.AppendLine("       END                                 AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Rif						AS Rif, ")
            loComandoSeleccionar.AppendLine("       CAST(@ldFecha_Desde AS DATE)        AS Fecha_De,")
            loComandoSeleccionar.AppendLine("       CAST(@ldFecha_Hasta AS DATE)        AS Fecha_Hasta")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("	JOIN	Cuentas_Pagar AS Documentos ")
            loComandoSeleccionar.AppendLine("		ON	Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loComandoSeleccionar.AppendLine("	JOIN	Retenciones_Documentos")
            loComandoSeleccionar.AppendLine("		ON	Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("	JOIN	Proveedores ")
            loComandoSeleccionar.AppendLine("		ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Retenciones")
            loComandoSeleccionar.AppendLine("		ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'RETIVA' ")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")

            loComandoSeleccionar.AppendLine("UPDATE #tabRetenciones")
            loComandoSeleccionar.AppendLine("SET #tabRetenciones.Rif = (SELECT CASE WHEN (ISNULL((SELECT Cod_Ter")
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
            loComandoSeleccionar.AppendLine("                                   END)")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT #tabRetenciones.Numero_Comprobante,")
            loComandoSeleccionar.AppendLine("       #tabRetenciones.Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("       #tabRetenciones.Numero_FacturaDoc,")
            'loComandoSeleccionar.AppendLine("       #tabRetenciones.Monto_Documento,")
            loComandoSeleccionar.AppendLine("       CASE WHEN #tabRetenciones.Fecha_Doc < '20180820' AND #tabRetenciones.Fecha_Doc > '20180731'")
            loComandoSeleccionar.AppendLine("            THEN (#tabRetenciones.Monto_Documento/100000)")
            loComandoSeleccionar.AppendLine("            ELSE #tabRetenciones.Monto_Documento")
            loComandoSeleccionar.AppendLine("       END AS Monto_Documento,")
            'loComandoSeleccionar.AppendLine("       #tabRetenciones.Monto_ExentoDoc,")
            loComandoSeleccionar.AppendLine("       CASE WHEN #tabRetenciones.Fecha_Doc < '20180820' AND #tabRetenciones.Fecha_Doc > '20180731'")
            loComandoSeleccionar.AppendLine("            THEN (#tabRetenciones.Monto_ExentoDoc/100000)")
            loComandoSeleccionar.AppendLine("            ELSE #tabRetenciones.Monto_ExentoDoc")
            loComandoSeleccionar.AppendLine("       END AS Monto_ExentoDoc,")
            'loComandoSeleccionar.AppendLine("       #tabRetenciones.Monto_BaseDoc,")
            loComandoSeleccionar.AppendLine("       CASE WHEN #tabRetenciones.Fecha_Doc < '20180820' AND #tabRetenciones.Fecha_Doc > '20180731'")
            loComandoSeleccionar.AppendLine("            THEN (#tabRetenciones.Monto_BaseDoc/100000)")
            loComandoSeleccionar.AppendLine("            ELSE #tabRetenciones.Monto_BaseDoc")
            loComandoSeleccionar.AppendLine("       END AS Monto_BaseDoc,")
            loComandoSeleccionar.AppendLine("       #tabRetenciones.Porcentaje_ImpuestoDoc,")
            loComandoSeleccionar.AppendLine("       #tabRetenciones.Fecha_Factura,")
            'loComandoSeleccionar.AppendLine("       #tabRetenciones.Base_Retencion,")
            loComandoSeleccionar.AppendLine("       CASE WHEN #tabRetenciones.Fecha_Doc < '20180820' AND #tabRetenciones.Fecha_Doc > '20180731'")
            loComandoSeleccionar.AppendLine("            THEN (#tabRetenciones.Base_Retencion/100000)")
            loComandoSeleccionar.AppendLine("            ELSE #tabRetenciones.Base_Retencion")
            loComandoSeleccionar.AppendLine("       END AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("       #tabRetenciones.Porcentaje_Retenido,")
            'loComandoSeleccionar.AppendLine("       #tabRetenciones.Monto_Retenido,")
            loComandoSeleccionar.AppendLine("       CASE WHEN #tabRetenciones.Fecha_Doc < '20180820' AND #tabRetenciones.Fecha_Doc > '20180731'")
            loComandoSeleccionar.AppendLine("            THEN (#tabRetenciones.Monto_Retenido/100000)")
            loComandoSeleccionar.AppendLine("            ELSE #tabRetenciones.Monto_Retenido")
            loComandoSeleccionar.AppendLine("       END AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("       #tabRetenciones.Nom_Pro,")
            loComandoSeleccionar.AppendLine("       #tabRetenciones.Rif,")
            loComandoSeleccionar.AppendLine("       #tabRetenciones.Fecha_De,")
            loComandoSeleccionar.AppendLine("       #tabRetenciones.Fecha_Hasta")
            loComandoSeleccionar.AppendLine("FROM #tabRetenciones")
            loComandoSeleccionar.AppendLine("ORDER BY Numero_Comprobante ASC")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DROP TABLE #tabRetenciones")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rResumenRetencion_IVAProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTRF_rResumenRetencion_IVAProveedores.ReportSource = loObjetoReporte


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
' RJG: 09/04/15: Codigo inicial, a partir de TRF_rResumenRetencion_IVAProveedores.                    '
'-------------------------------------------------------------------------------------------'
