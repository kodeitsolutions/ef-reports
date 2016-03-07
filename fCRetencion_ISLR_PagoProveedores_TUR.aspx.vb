'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCRetencion_ISLR_PagoProveedores_TUR"
'-------------------------------------------------------------------------------------------'
Partial Class fCRetencion_ISLR_PagoProveedores_TUR
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT          Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loConsulta.AppendLine("                Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loConsulta.AppendLine("                Cuentas_Pagar.Doc_Ori				AS Numero_Pago,")
            loConsulta.AppendLine("                Facturas_Retenidas.Factura		    AS Numero_Factura,")
            loConsulta.AppendLine("                Facturas_Retenidas.Control			AS Control_Factura,")
            loConsulta.AppendLine("                Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loConsulta.AppendLine("                Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loConsulta.AppendLine("                Renglones_Pagos.Mon_Net				AS Monto_Documento,")
            loConsulta.AppendLine("                Renglones_Pagos.Mon_Abo				AS Monto_Abonado,")
            loConsulta.AppendLine("                Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loConsulta.AppendLine("                Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loConsulta.AppendLine("                Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loConsulta.AppendLine("                RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loConsulta.AppendLine("                Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loConsulta.AppendLine("                Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loConsulta.AppendLine("                Proveedores.Nom_Pro					AS Nom_Pro,")
            loConsulta.AppendLine("                Proveedores.Rif						AS Rif,")
            loConsulta.AppendLine("                Proveedores.Nit						AS Nit,")
            loConsulta.AppendLine("                Proveedores.Dir_Fis					AS Direccion")
            loConsulta.AppendLine("FROM			Cuentas_Pagar")
            loConsulta.AppendLine("        JOIN    Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("        JOIN    Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loConsulta.AppendLine("        	AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loConsulta.AppendLine("        	AND Retenciones_Documentos.clase = 'ISLR'")
            loConsulta.AppendLine("        JOIN    Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loConsulta.AppendLine("        	AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine("        JOIN    Cuentas_Pagar Facturas_Retenidas")
            loConsulta.AppendLine("            ON  Facturas_Retenidas.Documento = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine("            AND Facturas_Retenidas.Cod_Tip = Retenciones_Documentos.Cod_Tip")
            loConsulta.AppendLine("        JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loConsulta.AppendLine("            AND	Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine("            AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")
            loConsulta.AppendLine("            AND " & goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
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

            '--------------------------------------------------'
            ' Carga la imagen del logo en curReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCRetencion_ISLR_PagoProveedores_TUR", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCRetencion_ISLR_PagoProveedores_TUR.ReportSource = loObjetoReporte


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
' RJG: 27/02/15: Código inicial, a partir de rCRetencion_ISLRProveedores_TUR.               '
'-------------------------------------------------------------------------------------------'
