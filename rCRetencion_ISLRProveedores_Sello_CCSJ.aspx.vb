'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCRetencion_ISLRProveedores_Sello_CCSJ"
'-------------------------------------------------------------------------------------------'
Partial Class rCRetencion_ISLRProveedores_Sello_CCSJ
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim loConsulta As New StringBuilder()


            loConsulta.AppendLine("")
            loConsulta.AppendLine("DECLARE @lcDocumento CHAR(10);")
            loConsulta.AppendLine("SET @lcDocumento= (  SELECT TOP 1 documento ")
            loConsulta.AppendLine("                     FROM pagos ")
            loConsulta.AppendLine("                     WHERE " &  cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("                   ); ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT		Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loConsulta.AppendLine("				Cuentas_Pagar.Doc_Ori				AS Numero_Pago,")
            loConsulta.AppendLine("				Origenes.Factura					AS Numero_Proveedor,")
            loConsulta.AppendLine("				Origenes.Control					AS Control_Proveedor,")
            loConsulta.AppendLine("				Origenes.Fec_Ini                    AS Fecha_Proveedor,")
            loConsulta.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loConsulta.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loConsulta.AppendLine("				Renglones_Pagos.Mon_Net				AS Monto_Documento,")
            loConsulta.AppendLine("				Renglones_Pagos.Mon_Abo				AS Monto_Abonado,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loConsulta.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loConsulta.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loConsulta.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loConsulta.AppendLine("				Proveedores.Rif						AS Rif,")
            loConsulta.AppendLine("				Proveedores.Nit						AS Nit,")
            loConsulta.AppendLine("				Proveedores.Dir_Fis					AS Direccion,")
            loConsulta.AppendLine("				Pagos.Recibo                        AS Recibo_Documento,")
            loConsulta.AppendLine("				Pagos.Control                       AS Control_Documento")
            loConsulta.AppendLine("FROM			Cuentas_Pagar")
            loConsulta.AppendLine("		JOIN	Pagos")
            loConsulta.AppendLine("		    ON  Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		    AND	Pagos.Documento = @lcDocumento")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.clase = 'ISLR'")
            loConsulta.AppendLine("		JOIN	Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loConsulta.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("		JOIN    Cuentas_Pagar AS Origenes ON Origenes.Cod_Tip = Retenciones_Documentos.Cod_Tip")
            loConsulta.AppendLine("			AND Origenes.Documento = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine("WHERE		Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")

            loConsulta.AppendLine("")
            loConsulta.AppendLine("UNION ALL		")
            loConsulta.AppendLine("")

            loConsulta.AppendLine("SELECT		Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loConsulta.AppendLine("				''									AS Numero_Pago,")
            loConsulta.AppendLine("				''									AS Numero_Proveedor,")
            loConsulta.AppendLine("				''									AS Control_Proveedor,")
            loConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Proveedor,")
            loConsulta.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loConsulta.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loConsulta.AppendLine("				Documentos.Mon_Net					AS Monto_Documento,")
            loConsulta.AppendLine("				Documentos.Mon_Net					AS Monto_Abonado,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loConsulta.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loConsulta.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loConsulta.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loConsulta.AppendLine("				Proveedores.Rif						AS Rif,")
            loConsulta.AppendLine("				Proveedores.Nit						AS Nit,")
            loConsulta.AppendLine("				Proveedores.Dir_Fis					AS Direccion,")
            loConsulta.AppendLine("				''                                  AS Recibo_Documento,")
            loConsulta.AppendLine("				''                                  AS Control_Documento")
            loConsulta.AppendLine("FROM			Cuentas_Pagar")
            loConsulta.AppendLine("     JOIN    renglones_pagos AS Retenciones_Pagadas ")
            loConsulta.AppendLine("         ON  Retenciones_Pagadas.cod_tip = cuentas_pagar.cod_tip")
            loConsulta.AppendLine("         AND Retenciones_Pagadas.doc_ori = cuentas_pagar.documento")
            loConsulta.AppendLine("         AND Retenciones_Pagadas.documento = @lcDocumento")
            loConsulta.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("	LEFT JOIN	Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE		Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("rCRetencion_ISLRProveedores_Sello_CCSJ", laDatosReporte)
            'loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCRetencion_ISLRProveedores_Sello_CCSJ", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCRetencion_ISLRProveedores_Sello_CCSJ.ReportSource = loObjetoReporte


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
' RJG: 28/03/14: Codigo inicial, a partir de rCRetencion_ISLRProveedores_CCSJ_002.      	'
'-------------------------------------------------------------------------------------------'
