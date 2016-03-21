'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rResumenRetencion_IVAProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rResumenRetencion_IVAProveedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcAño As String = CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Year)
            Dim lcMes As String = CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Month)

            If lcMes < 10 Then
                lcMes = "0" & lcMes
            End If

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine(" SELECT '" & lcAño & "' + '" & lcMes & "' + Retenciones_Documentos.Num_Com		AS Numero_Comprobante,")
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
            loComandoSeleccionar.AppendLine("       CAST(" & lcParametro0Desde & " AS DATE)  AS Fecha_De,")
            loComandoSeleccionar.AppendLine("       CAST(" & lcParametro0Hasta & " AS DATE)  AS Fecha_Hasta")
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

            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Fec_Ini BETWEEN  " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 	AND " & lcParametro0Hasta)


            loComandoSeleccionar.AppendLine("UNION ALL	")

            loComandoSeleccionar.AppendLine("SELECT	'" & lcAño & "' + '" & lcMes & "'  + Retenciones_Documentos.Num_Com		AS Numero_Comprobante,")
            loComandoSeleccionar.AppendLine("		Documentos.Control					AS Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Factura					AS Numero_FacturaDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Mon_Net					AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("		Documentos.Mon_Exe					AS Monto_ExentoDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Mon_Bas1					AS Monto_BaseDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,")
            loComandoSeleccionar.AppendLine("		Documentos.Fec_Ini					AS Fecha_Factura,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Rif						AS Rif, ")
            loComandoSeleccionar.AppendLine("       CAST(" & lcParametro0Desde & " AS DATE)  AS Fecha_De,")
            loComandoSeleccionar.AppendLine("       CAST(" & lcParametro0Hasta & " AS DATE)  AS Fecha_Hasta")
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

            loComandoSeleccionar.AppendLine("	AND Cuentas_Pagar.Fec_Ini BETWEEN  " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine(" 	AND " & lcParametro0Hasta)

            loComandoSeleccionar.AppendLine("ORDER BY Numero_Comprobante ASC")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rResumenRetencion_IVAProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCGS_rResumenRetencion_IVAProveedores.ReportSource = loObjetoReporte


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
' RJG: 09/04/15: Codigo inicial, a partir de CGS_rResumenRetencion_IVAProveedores.                    '
'-------------------------------------------------------------------------------------------'
