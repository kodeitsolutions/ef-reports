Imports System.Data
Partial Class fCRetencion_ISLR_PagoProveedores_GPV
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT          Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("                Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("                Cuentas_Pagar.Doc_Ori				AS Numero_Pago,")
            loComandoSeleccionar.AppendLine("                Facturas_Retenidas.Factura		    AS Numero_Factura,")
            loComandoSeleccionar.AppendLine("                Facturas_Retenidas.Control			AS Control_Factura,")
            loComandoSeleccionar.AppendLine("                Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("                Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("                Renglones_Pagos.Mon_Net				AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("                Renglones_Pagos.Mon_Abo				AS Monto_Abonado,")
            loComandoSeleccionar.AppendLine("                Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("                Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("                Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loComandoSeleccionar.AppendLine("                RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loComandoSeleccionar.AppendLine("                Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("                Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("                Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("                Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("                Proveedores.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("                Proveedores.Dir_Fis					AS Direccion")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("        JOIN    Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("        JOIN    Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("        	AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loComandoSeleccionar.AppendLine("        	AND Retenciones_Documentos.clase = 'ISLR'")
            loComandoSeleccionar.AppendLine("        JOIN    Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("        	AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("        JOIN    Cuentas_Pagar Facturas_Retenidas")
            loComandoSeleccionar.AppendLine("            ON  Facturas_Retenidas.Documento = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("            AND Facturas_Retenidas.Cod_Tip = Retenciones_Documentos.Cod_Tip")
            loComandoSeleccionar.AppendLine("        JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("            AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("            AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")
            loComandoSeleccionar.AppendLine("            AND " & goFormatos.pcCondicionPrincipal)
            loComandoSeleccionar.AppendLine("")
            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")


            '--------------------------------------------------'
            ' Carga la imagen del logo en curReportes          '
            '--------------------------------------------------'
            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCRetencion_ISLR_PagoProveedores_GPV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCRetencion_ISLR_PagoProveedores_GPV.ReportSource = loObjetoReporte

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
' EAG: 18/09/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
