'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCRetencion_ISLRProveedores_SVT"
'-------------------------------------------------------------------------------------------'
Partial Class rCRetencion_ISLRProveedores_SVT
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
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim lcConsulta As New StringBuilder()




            lcConsulta.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            lcConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            lcConsulta.AppendLine("				Cuentas_Pagar.Doc_Ori				AS Numero_Pago,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            lcConsulta.AppendLine("				Renglones_Pagos.Factura		        AS Factura_Documento,")
            lcConsulta.AppendLine("				Renglones_Pagos.Mon_Net				AS Monto_Documento,")
            lcConsulta.AppendLine("				Renglones_Pagos.Mon_Abo				AS Monto_Abonado,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            lcConsulta.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            lcConsulta.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            lcConsulta.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            lcConsulta.AppendLine("				Proveedores.Rif						AS Rif,")
            lcConsulta.AppendLine("				Proveedores.Nit						AS Nit,")
            lcConsulta.AppendLine("				Proveedores.Dir_Fis					AS Direccion")
            lcConsulta.AppendLine("FROM			Cuentas_Pagar")
            lcConsulta.AppendLine("		JOIN	Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            lcConsulta.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            lcConsulta.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            lcConsulta.AppendLine("			AND Retenciones_Documentos.clase = 'ISLR'")
            lcConsulta.AppendLine("		JOIN	Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            lcConsulta.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            lcConsulta.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            lcConsulta.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            lcConsulta.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            lcConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            lcConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")

            lcConsulta.AppendLine("           AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            lcConsulta.AppendLine("         		AND " & lcParametro0Hasta)
            lcConsulta.AppendLine("           AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            lcConsulta.AppendLine("         		AND " & lcParametro1Hasta)
            lcConsulta.AppendLine("           AND Pagos.Cod_Mon BETWEEN " & lcParametro2Desde)
            lcConsulta.AppendLine("         		AND " & lcParametro2Hasta)
            lcConsulta.AppendLine("           AND Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            lcConsulta.AppendLine("         		AND " & lcParametro3Hasta)

            lcConsulta.AppendLine("UNION ALL		")

            lcConsulta.AppendLine("SELECT			Retenciones_Documentos.Tip_Ori		AS Tipo_Origen,")
            lcConsulta.AppendLine("				Ordenes_Pagos.Fec_Ini				AS Fecha_Retencion,")
            lcConsulta.AppendLine("				''									AS Numero_Pago,")
            lcConsulta.AppendLine("				'ORD/PAG'							AS Tipo_Documento,")
            lcConsulta.AppendLine("				Ordenes_Pagos.Documento				AS Numero_Documento,")
            lcConsulta.AppendLine("				Ordenes_Pagos.Documento		        AS Factura_Documento,")
            lcConsulta.AppendLine("				Ordenes_Pagos.Mon_Net				AS Monto_Documento,")
            lcConsulta.AppendLine("				Ordenes_Pagos.Mon_Net				AS Monto_Abonado,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            lcConsulta.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            lcConsulta.AppendLine("				Ordenes_Pagos.Cod_Pro				AS Cod_Pro,")
            lcConsulta.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            lcConsulta.AppendLine("				Proveedores.Rif						AS Rif,")
            lcConsulta.AppendLine("				Proveedores.Nit						AS Nit,")
            lcConsulta.AppendLine("				Proveedores.Dir_Fis					AS Direccion")
            lcConsulta.AppendLine("FROM			Retenciones_Documentos")
            lcConsulta.AppendLine("	JOIN		Ordenes_Pagos ON Ordenes_Pagos.Documento = Retenciones_Documentos.documento")
            lcConsulta.AppendLine("	JOIN		Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            lcConsulta.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            lcConsulta.AppendLine("WHERE		Ordenes_Pagos.Status = 'Confirmado'")
            lcConsulta.AppendLine("			AND	Retenciones_Documentos.Tip_Ori	= 'Ordenes_Pagos'")
            lcConsulta.AppendLine("			AND Retenciones_Documentos.clase	= 'ISLR'")

            lcConsulta.AppendLine("           AND Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde)
            lcConsulta.AppendLine("         		AND " & lcParametro0Hasta)
            lcConsulta.AppendLine("           AND Ordenes_Pagos.Cod_Pro BETWEEN " & lcParametro1Desde)
            lcConsulta.AppendLine("         		AND " & lcParametro1Hasta)
            lcConsulta.AppendLine("           AND Ordenes_Pagos.Cod_Mon BETWEEN " & lcParametro2Desde)
            lcConsulta.AppendLine("         		AND " & lcParametro2Hasta)
            lcConsulta.AppendLine("           AND Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            lcConsulta.AppendLine("         		AND " & lcParametro3Hasta)


            lcConsulta.AppendLine("UNION ALL		")

            lcConsulta.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            lcConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            lcConsulta.AppendLine("				''									AS Numero_Pago,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            lcConsulta.AppendLine("				Documentos.Factura		            AS Factura_Documento,")
            lcConsulta.AppendLine("				Documentos.Mon_Net					AS Monto_Documento,")
            lcConsulta.AppendLine("				Documentos.Mon_Net					AS Monto_Abonado,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            lcConsulta.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            lcConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            lcConsulta.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            lcConsulta.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            lcConsulta.AppendLine("				Proveedores.Rif						AS Rif,")
            lcConsulta.AppendLine("				Proveedores.Nit						AS Nit,")
            lcConsulta.AppendLine("				Proveedores.Dir_Fis					AS Direccion")
            lcConsulta.AppendLine("FROM			Cuentas_Pagar")
            lcConsulta.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            lcConsulta.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            lcConsulta.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            lcConsulta.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            lcConsulta.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            lcConsulta.AppendLine("	LEFT JOIN	Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            lcConsulta.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            lcConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            lcConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            lcConsulta.AppendLine("       	    AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            lcConsulta.AppendLine("       	  		AND " & lcParametro0Hasta)
            lcConsulta.AppendLine("       	    AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            lcConsulta.AppendLine("       	  		AND " & lcParametro1Hasta)
            lcConsulta.AppendLine("       	    AND Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro2Desde)
            lcConsulta.AppendLine("       	  		AND " & lcParametro2Hasta)
            lcConsulta.AppendLine("       	    AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
            lcConsulta.AppendLine("       	  		AND " & lcParametro3Hasta)

            lcConsulta.AppendLine("ORDER BY " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(lcConsulta.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCRetencion_ISLRProveedores_SVT", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrrCRetencion_ISLRProveedores_SVT.ReportSource = loObjetoReporte


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
' EAG: 07/10/15: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
