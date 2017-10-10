'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "RAL_rResumenRetencion_ISLRProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class RAL_rResumenRetencion_ISLRProveedores
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

            loComandoSeleccionar.AppendLine("DECLARE @lnCero AS DECIMAL(28, 10) = CAST(0 AS DECIMAL(28, 10));")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio AS NVARCHAR(30) = N''; ")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Desde AS DATETIME = " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("DECLARE @ldFecha_Hasta AS DATETIME = " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Desde AS VARCHAR(10) = " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("DECLARE @lcCodPro_Hasta AS VARCHAR(10) = " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT	Documentos.Factura		            AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("       Documentos.Fec_Ini					AS Fec_Factura,")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("       Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("       Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("       CAST(@ldFecha_Desde AS DATE)        AS Fecha_De,")
            loComandoSeleccionar.AppendLine("       CAST(@ldFecha_Hasta AS DATE)        AS Fecha_Hasta,")
            loComandoSeleccionar.AppendLine("		@lcVacio							AS Fec_Recibo,")
            loComandoSeleccionar.AppendLine("		@lnCero								AS Mon_Ret,")
            loComandoSeleccionar.AppendLine("		@lcVacio							AS Porcentaje,")
            loComandoSeleccionar.AppendLine("		@lcVacio							AS Rif_Tra,")
            loComandoSeleccionar.AppendLine("		@lcVacio							AS Nom_Tra,")
            loComandoSeleccionar.AppendLine("       'PRO'								AS Tipo	")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("   JOIN Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("       AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loComandoSeleccionar.AppendLine("   JOIN Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("        AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("   JOIN Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("   AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT Renglones_Pagos.Factura		        AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("		(SELECT Fec_Ini")
            loComandoSeleccionar.AppendLine("		FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("		WHERE Factura = Renglones_Pagos.Factura")
            loComandoSeleccionar.AppendLine("		    AND Cod_Pro = Proveedores.Cod_Pro)AS Fec_Factura,")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("       Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("       Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("       Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("       Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("       CAST(@ldFecha_Desde AS DATE)        AS Fecha_De,")
            loComandoSeleccionar.AppendLine("       CAST(@ldFecha_Hasta AS DATE)        AS Fecha_Hasta,")
            loComandoSeleccionar.AppendLine("		@lcVacio							AS Fec_Recibo,")
            loComandoSeleccionar.AppendLine("		@lnCero								AS Mon_Ret,")
            loComandoSeleccionar.AppendLine("		@lcVacio							AS Porcentaje,")
            loComandoSeleccionar.AppendLine("		@lcVacio							AS Rif_Tra,")
            loComandoSeleccionar.AppendLine("		@lcVacio							AS Nom_Tra,")
            loComandoSeleccionar.AppendLine("       'PRO'								AS Tipo	")
            loComandoSeleccionar.AppendLine("FROM Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("   JOIN Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("   JOIN Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("       AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loComandoSeleccionar.AppendLine("       AND Retenciones_Documentos.clase = 'ISLR'")
            loComandoSeleccionar.AppendLine("   JOIN Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("       AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("   JOIN Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("   LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("   AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("   AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("   AND Cuentas_Pagar.Cod_Pro BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("UNION ALL")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT @lcVacio		                AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("       @lcVacio					    AS Fec_Factura,")
            loComandoSeleccionar.AppendLine("       @lnCero		                    AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("       @lnCero		                    AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("       @lnCero		                    AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("       @lcVacio				        AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("       @lcVacio					    AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("       @lcVacio						AS Rif,")
            loComandoSeleccionar.AppendLine("       CAST(@ldFecha_Desde AS DATE)    AS Fecha_De,")
            loComandoSeleccionar.AppendLine("       CAST(@ldFecha_Hasta AS DATE)    AS Fecha_Hasta,")
            loComandoSeleccionar.AppendLine("       Renglones_Recibos.Fec_Ini		AS Fec_Recibo,")
            loComandoSeleccionar.AppendLine("       Renglones_Recibos.Mon_Net		AS Mon_Ret,")
            loComandoSeleccionar.AppendLine("       Renglones_Recibos.val_car		AS Porcentaje,")
            loComandoSeleccionar.AppendLine("       Trabajadores.Rif    			AS Rif_Tra,")
            loComandoSeleccionar.AppendLine("       Trabajadores.Nom_Tra			AS Nom_Tra,")
            loComandoSeleccionar.AppendLine("       'TRA'							AS Tipo	")
            loComandoSeleccionar.AppendLine("FROM Renglones_Recibos")
            loComandoSeleccionar.AppendLine("   JOIN Recibos ON Renglones_Recibos.documento = Recibos.documento")
            loComandoSeleccionar.AppendLine("       AND Renglones_Recibos.cod_con IN ('R005', 'R405')")
            loComandoSeleccionar.AppendLine("   JOIN Nominas ON Recibos.Doc_Ori = Nominas.Documento")
            loComandoSeleccionar.AppendLine("   JOIN Trabajadores ON Recibos.cod_tra = Trabajadores.cod_tra")
            loComandoSeleccionar.AppendLine("WHERE Nominas.Fec_Ini BETWEEN @ldFecha_Desde AND @ldFecha_Hasta")
            loComandoSeleccionar.AppendLine("   AND Recibos.Cod_Tra BETWEEN @lcCodPro_Desde AND @lcCodPro_Hasta")
            'loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("ORDER BY Fec_Factura ASC")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("RAL_rResumenRetencion_ISLRProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRAL_rResumenRetencion_ISLRProveedores.ReportSource = loObjetoReporte


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
' RJG: 09/04/15: Codigo inicial, a partir de RAL_rResumenRetencion_ISLRProveedores.                    '
'-------------------------------------------------------------------------------------------'
