'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "TRF_rResumenRetencion_ISLRProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class TRF_rResumenRetencion_ISLRProveedores
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

            loComandoSeleccionar.AppendLine("DECLARE @lnCero AS DECIMAL(28, 10);")
            loComandoSeleccionar.AppendLine("SET @lnCero = CAST(0 AS DECIMAL(28, 10));")
            loComandoSeleccionar.AppendLine("DECLARE @lcVacio AS NVARCHAR(30); ")
            loComandoSeleccionar.AppendLine("SET @lcVacio = N''; ")
            loComandoSeleccionar.AppendLine("")

            loComandoSeleccionar.AppendLine("SELECT			Documentos.Factura		            AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("               Documentos.Fec_Ini					AS Fec_Factura,")
            loComandoSeleccionar.AppendLine("            	Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("            	Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("            	Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("            	Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("            	Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("            	Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("               CAST(" & lcParametro0Desde & " AS DATE)  AS Fecha_De,")
            loComandoSeleccionar.AppendLine("               CAST(" & lcParametro0Hasta & " AS DATE)  AS Fecha_Hasta,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							AS Fec_Recibo,")
            loComandoSeleccionar.AppendLine("			    @lnCero								AS Mon_Ret,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							AS Porcentaje,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							AS Rif_Tra,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							AS Nom_Tra,")
            loComandoSeleccionar.AppendLine("               'PRO'									AS Tipo	")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("        JOIN	Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("            AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loComandoSeleccionar.AppendLine("        JOIN	Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("            AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("        JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("    LEFT JOIN	Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("            AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("            AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("                AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("                AND " & lcParametro1Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		")

            loComandoSeleccionar.AppendLine("SELECT			Ordenes_Pagos.Documento             AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("               Ordenes_Pagos.Fec_Ini				AS Fec_Factura,")
            loComandoSeleccionar.AppendLine("               Renglones_OPagos.Mon_Net                  AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("               ISNULL((SELECT CAST(R_OP.Comentario AS DECIMAL (28,2))")
            loComandoSeleccionar.AppendLine("                       FROM Renglones_OPagos AS R_OP")
            loComandoSeleccionar.AppendLine("                       WHERE R_OP.Cod_Con = 'P0007'")
            loComandoSeleccionar.AppendLine("                           AND R_OP.Documento = Renglones_OPagos.Documento")
            loComandoSeleccionar.AppendLine("               ),@lnCero)                                AS Porcentaje_Retencion,")
            loComandoSeleccionar.AppendLine("               ISNULL((SELECT R_OP.Mon_Net")
            loComandoSeleccionar.AppendLine("                       FROM Renglones_OPagos AS R_OP")
            loComandoSeleccionar.AppendLine("                       WHERE R_OP.Cod_Con = 'P0007'")
            loComandoSeleccionar.AppendLine("                           AND R_OP.Documento = Renglones_OPagos.Documento")
            loComandoSeleccionar.AppendLine("               ),@lnCero)                              AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("            	Ordenes_Pagos.Cod_Pro				    AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("            	Proveedores.Nom_Pro					    AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("            	Proveedores.Rif						    AS Rif,")
            loComandoSeleccionar.AppendLine("               CAST(" & lcParametro0Desde & " AS DATE) AS Fecha_De,")
            loComandoSeleccionar.AppendLine("               CAST(" & lcParametro0Hasta & " AS DATE) AS Fecha_Hasta,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							    AS Fec_Recibo,")
            loComandoSeleccionar.AppendLine("               @lnCero                                 AS Mon_Ret,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							    AS Porcentaje,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							    AS Rif_Tra,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							    AS Nom_Tra,")
            loComandoSeleccionar.AppendLine("               'TRA'									AS Tipo	")
            loComandoSeleccionar.AppendLine("FROM			Ordenes_Pagos")
            loComandoSeleccionar.AppendLine("        JOIN	Renglones_OPagos ON Ordenes_Pagos.Documento = Renglones_OPagos.Documento")
            loComandoSeleccionar.AppendLine("        JOIN	Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE		Renglones_OPagos.Cod_Con = 'E0101'")
            loComandoSeleccionar.AppendLine("            AND Ordenes_Pagos.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("            AND Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("                AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("                AND " & lcParametro1Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		")

            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Documento                 AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("               Cuentas_Pagar.Fec_Ini				    AS Fec_Factura,")
            loComandoSeleccionar.AppendLine("               Campos_Extras.Doble1                    AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("               Campos_Extras.Doble2                    AS Porcentaje_Retencion,")
            loComandoSeleccionar.AppendLine("               Cuentas_Pagar.Mon_Net                   AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("            	Cuentas_Pagar.Cod_Pro				    AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("            	Proveedores.Nom_Pro					    AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("            	Proveedores.Rif						    AS Rif,")
            loComandoSeleccionar.AppendLine("               CAST(" & lcParametro0Desde & " AS DATE) AS Fecha_De,")
            loComandoSeleccionar.AppendLine("               CAST(" & lcParametro0Hasta & " AS DATE) AS Fecha_Hasta,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							    AS Fec_Recibo,")
            loComandoSeleccionar.AppendLine("               @lnCero                                 AS Mon_Ret,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							    AS Porcentaje,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							    AS Rif_Tra,")
            loComandoSeleccionar.AppendLine("			    @lcVacio							    AS Nom_Tra,")
            loComandoSeleccionar.AppendLine("               'TRA'									AS Tipo	")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("        JOIN Campos_Extras ON Cuentas_Pagar.Documento = Campos_Extras.Cod_Reg")
            loComandoSeleccionar.AppendLine("           AND Origen = 'Cuentas_Pagar' AND Clase = 'ISLR'")
            loComandoSeleccionar.AppendLine("        JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("            AND Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("            AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("                AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("                AND " & lcParametro1Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		")

            loComandoSeleccionar.AppendLine("SELECT			@lcVacio		                AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("               @lcVacio					    AS Fec_Factura,")
            loComandoSeleccionar.AppendLine("            	@lnCero		                    AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("            	@lnCero		                    AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("            	@lnCero		                    AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("            	@lcVacio				        AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("            	@lcVacio					    AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("            	@lcVacio						AS Rif,")
            loComandoSeleccionar.AppendLine("               CAST(" & lcParametro0Desde & " AS DATE)  AS Fecha_De,")
            loComandoSeleccionar.AppendLine("               CAST(" & lcParametro0Hasta & " AS DATE)  AS Fecha_Hasta,")
            loComandoSeleccionar.AppendLine("               Renglones_Recibos.Fec_Ini		AS Fec_Recibo,")
            loComandoSeleccionar.AppendLine("               Renglones_Recibos.Mon_Net		AS Mon_Ret,")
            loComandoSeleccionar.AppendLine("               Renglones_Recibos.val_car		AS Porcentaje,")
            loComandoSeleccionar.AppendLine("               Trabajadores.Rif    			AS Rif_Tra,")
            loComandoSeleccionar.AppendLine("               Trabajadores.Nom_Tra			AS Nom_Tra,")
            loComandoSeleccionar.AppendLine("               'TRA'									AS Tipo	")
            loComandoSeleccionar.AppendLine("FROM Renglones_Recibos")
            loComandoSeleccionar.AppendLine("   JOIN Recibos ON Renglones_Recibos.documento = Recibos.documento")
            loComandoSeleccionar.AppendLine("       AND Renglones_Recibos.cod_con IN ('R005', 'R405')")
            loComandoSeleccionar.AppendLine("   JOIN Nominas ON Recibos.Doc_Ori = Nominas.Documento")
            loComandoSeleccionar.AppendLine("JOIN Trabajadores ON Recibos.cod_tra = Trabajadores.cod_tra")
            loComandoSeleccionar.AppendLine("WHERE     Nominas.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("                AND " & lcParametro0Hasta)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("TRF_rResumenRetencion_ISLRProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTRF_rResumenRetencion_ISLRProveedores.ReportSource = loObjetoReporte


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
' RJG: 09/04/15: Codigo inicial, a partir de TRF_rResumenRetencion_ISLRProveedores.                    '
'-------------------------------------------------------------------------------------------'
