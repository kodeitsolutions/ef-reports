'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCRetencion_IVAProveedores_GPV"
'-------------------------------------------------------------------------------------------'
Partial Class rCRetencion_IVAProveedores_GPV
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

            Dim lcAño As String = "'" & Strings.Right("0000" & CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Year), 4) & "'"
            Dim lcMes As String = "'" & Strings.Right("00" & CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Month), 2) & "'"
            Dim loConsulta As New StringBuilder()


            loConsulta.AppendLine("SELECT		Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loConsulta.AppendLine("				Retenciones_Documentos.Num_Com		AS Numero_Comprobante,")
            loConsulta.AppendLine("				Cuentas_Pagar.Documento				AS Numero_DocumentoRet,")
            loConsulta.AppendLine("				Cuentas_Pagar.Control				AS Numero_ControlRet,")
            loConsulta.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loConsulta.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loConsulta.AppendLine("				Documentos.Factura		            AS Numero_FacturaDoc,")
            loConsulta.AppendLine("				Documentos.Control					AS Numero_ControlDoc,")
            loConsulta.AppendLine("				Renglones_Pagos.Mon_Net				AS Monto_Documento,")
            loConsulta.AppendLine("				Documentos.Mon_Exe					AS Monto_ExentoDoc,	")
            loConsulta.AppendLine("				Documentos.Mon_Bas1					AS Monto_BaseDoc,")
            loConsulta.AppendLine("				Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,")
            loConsulta.AppendLine("				Documentos.Mon_Imp1					AS Monto_ImpuestoDoc,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loConsulta.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loConsulta.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loConsulta.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loConsulta.AppendLine("				Proveedores.Rif						AS Rif,")
            loConsulta.AppendLine("				Proveedores.Nit						AS Nit,")
            loConsulta.AppendLine("				Proveedores.Dir_Fis					AS Direccion,")
            loConsulta.AppendLine("				" & lcAño & "						AS Anio,")
            loConsulta.AppendLine("				" & lcMes & "						AS Mes")
            loConsulta.AppendLine("FROM			Cuentas_Pagar")
            loConsulta.AppendLine("		JOIN	Pagos")
            loConsulta.AppendLine("			ON	Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos ")
            loConsulta.AppendLine("			ON	Retenciones_Documentos.Documento = Pagos.Documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.Documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.clase	= 'IMPUESTO'")
            loConsulta.AppendLine("		JOIN	Renglones_Pagos ")
            loConsulta.AppendLine("			ON	Renglones_Pagos.Documento = Pagos.Documento")
            loConsulta.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine("		LEFT JOIN	Cuentas_Pagar AS Documentos										")
            loConsulta.AppendLine("			ON	Documentos.Documento = Renglones_Pagos.Doc_Ori					")
            loConsulta.AppendLine("			AND	Documentos.Cod_Tip = Renglones_Pagos.Cod_Tip 					")
            loConsulta.AppendLine("		JOIN	Proveedores")
            loConsulta.AppendLine("			ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("		LEFT JOIN Retenciones")
            loConsulta.AppendLine("			ON	Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE		Cuentas_Pagar.Cod_Tip = 'RETIVA'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")

            loConsulta.AppendLine("           AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("         		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("           AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("         		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("           AND Pagos.Cod_Mon BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("         		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("           AND Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("         		AND " & lcParametro3Hasta)

            loConsulta.AppendLine("UNION ALL		")

            loConsulta.AppendLine("SELECT		Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loConsulta.AppendLine("				Retenciones_Documentos.Num_Com		AS Numero_Comprobante,				")
            loConsulta.AppendLine("				Cuentas_Pagar.Documento				AS Numero_DocumentoRet,				")
            loConsulta.AppendLine("				Cuentas_Pagar.Control				AS Numero_ControlRet,				")
            loConsulta.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loConsulta.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loConsulta.AppendLine("				Documentos.Factura		            AS Numero_FacturaDoc,")
            loConsulta.AppendLine("				Documentos.Control					AS Numero_ControlDoc,")
            loConsulta.AppendLine("				Documentos.Mon_Net					AS Monto_Documento,")
            loConsulta.AppendLine("				Documentos.Mon_Exe					AS Monto_ExentoDoc,					")
            loConsulta.AppendLine("				Documentos.Mon_Bas1					AS Monto_BaseDoc,					")
            loConsulta.AppendLine("				Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,			")
            loConsulta.AppendLine("				Documentos.Mon_Imp1					AS Monto_ImpuestoDoc,				")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loConsulta.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loConsulta.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loConsulta.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loConsulta.AppendLine("				Proveedores.Rif						AS Rif,")
            loConsulta.AppendLine("				Proveedores.Nit						AS Nit,")
            loConsulta.AppendLine("				Proveedores.Dir_Fis					AS Direccion,")
            loConsulta.AppendLine("				" & lcAño & "						AS Anio,")
            loConsulta.AppendLine("				" & lcMes & "						AS Mes")
            loConsulta.AppendLine("FROM			Cuentas_Pagar")
            loConsulta.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ")
            loConsulta.AppendLine("			ON	Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos")
            loConsulta.AppendLine("			ON	Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Proveedores ")
            loConsulta.AppendLine("			ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("		LEFT JOIN	Retenciones")
            loConsulta.AppendLine("			ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'RETIVA'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loConsulta.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")

            loConsulta.AppendLine("       	    AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("       	  		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("       	    AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("       	  		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("       	    AND Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("       	  		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("       	    AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("       	  		AND " & lcParametro3Hasta)

            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento & ", Fecha_Retencion ASC")



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCRetencion_IVAProveedores_GPV", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrrCRetencion_IVAProveedores_GPV.ReportSource = loObjetoReporte


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
' EAG: 09/09/15: Codigo inicial
'-------------------------------------------------------------------------------------------'
