'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rRDetallada_ISLR_Retenido"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rRDetallada_ISLR_Retenido
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Desde AS VARCHAR(15) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Hasta AS VARCHAR(15) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Documentos.Factura				AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("		Documentos.Control				AS Control,")
            loComandoSeleccionar.AppendLine("		Documentos.Mon_Net				AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini			AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("		Documentos.Mon_Net				AS Monto_Abonado,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Mon_Bas	AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Por_Ret	AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("		Retenciones.Cod_Ret				AS Codigo_Concepto,")
            loComandoSeleccionar.AppendLine("		Retenciones.Nom_Ret				AS Concepto,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Mon_Ret	AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Pro			AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro				AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("		Proveedores.Rif					AS Rif,")
            loComandoSeleccionar.AppendLine("		Proveedores.Dir_Fis				AS Direccion,")
            loComandoSeleccionar.AppendLine("       @ldFecha_Desde					AS Desde,")
            loComandoSeleccionar.AppendLine("       @ldFecha_Hasta					AS Hasta")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loComandoSeleccionar.AppendLine("	JOIN Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Facturas.Factura					AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("		Facturas.Control					AS Control,")
            loComandoSeleccionar.AppendLine("		Facturas.Mon_Net					AS Monto_Documento,	")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("		Renglones_Pagos.Mon_Abo				AS Monto_Abonado,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("		Retenciones.Cod_Ret					AS Codigo_Concepto,	")
            loComandoSeleccionar.AppendLine("		Retenciones.Nom_Ret					AS Concepto,")
            loComandoSeleccionar.AppendLine("		Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("		Cuentas_Pagar.Cod_Pro				AS Cod_Pro,	")
            loComandoSeleccionar.AppendLine("		Proveedores.Nom_Pro					AS Nom_Pro,	")
            loComandoSeleccionar.AppendLine("		Proveedores.Rif						AS Rif,	")
            loComandoSeleccionar.AppendLine("		Proveedores.Dir_Fis					AS Direccion,")
            loComandoSeleccionar.AppendLine("		@ldFecha_Desde						AS Desde,")
            loComandoSeleccionar.AppendLine("		@ldFecha_Hasta						AS Hasta")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("	JOIN Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("	JOIN Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.Doc_Des = Cuentas_Pagar.documento")
            loComandoSeleccionar.AppendLine("	JOIN Cuentas_Pagar AS Facturas ON Retenciones_Documentos.Doc_Ori = Facturas.Documento")
            loComandoSeleccionar.AppendLine("		AND Facturas.Cod_Tip = Retenciones_Documentos.Cla_Ori")
            loComandoSeleccionar.AppendLine("	JOIN Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("		AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("	JOIN Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("   AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("	AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            loComandoSeleccionar.AppendLine("ORDER BY Cod_Pro, Fecha_Retencion")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº 0) trae registros
            '-------------------------------------------------------------------------------------------------------

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rRDetallada_ISLR_Retenido", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCGS_rRDetallada_ISLR_Retenido.ReportSource = loObjetoReporte


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
' MAT:  07/06/11: Codigo inicial
'-------------------------------------------------------------------------------------------'
