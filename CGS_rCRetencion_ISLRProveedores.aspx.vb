'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rCRetencion_ISLRProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rCRetencion_ISLRProveedores
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Doc_Ori				AS Numero_Pago,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Control				AS Control_Documento,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Factura				AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Fec_Ini			    AS Fecha_Factura,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Mon_Net				AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Mon_Abo				AS Monto_Abonado,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loComandoSeleccionar.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Proveedores.Dir_Fis					AS Direccion,")
            loComandoSeleccionar.AppendLine("               " & lcParametro4Desde & " AS Agrupar")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("		JOIN	Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.clase = 'ISLR'")
            loComandoSeleccionar.AppendLine("		JOIN	Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("			AND	(Cuentas_Pagar.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("			    AND Cuentas_Pagar.Status IN (SELECT Documentos.status")
            loComandoSeleccionar.AppendLine("               FROM cuentas_pagar")
            loComandoSeleccionar.AppendLine("			    JOIN retenciones_documentos ON cuentas_pagar.documento = retenciones_documentos.doc_ori")
            loComandoSeleccionar.AppendLine("			        AND cuentas_pagar.cod_tip = 'FACT'")
            loComandoSeleccionar.AppendLine("			    JOIN cuentas_pagar AS Documentos ON retenciones_documentos.doc_des = Documentos.documento")
            loComandoSeleccionar.AppendLine("			        AND Documentos.cod_tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("			    WHERE cuentas_pagar.factura BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		    AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			        AND Documentos.doc_ori = retenciones_documentos.documento))")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")

            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Renglones_Pagos.Factura BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		")

            loComandoSeleccionar.AppendLine("SELECT			Retenciones_Documentos.Tip_Ori		AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				''									AS Numero_Pago,")
            loComandoSeleccionar.AppendLine("				'ORD/PAG'							AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Documento				AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("				''								    AS Control_Documento,")
            loComandoSeleccionar.AppendLine("				''									AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Fec_Ini               AS Fecha_Factura,")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Mon_Net				AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Mon_Net				AS Monto_Abonado,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loComandoSeleccionar.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Proveedores.Dir_Fis					AS Direccion,")
            loComandoSeleccionar.AppendLine("               " & lcParametro4Desde & " AS Agrupar")
            loComandoSeleccionar.AppendLine("FROM			Retenciones_Documentos")
            loComandoSeleccionar.AppendLine("	JOIN		Ordenes_Pagos ON Ordenes_Pagos.Documento = Retenciones_Documentos.documento")
            loComandoSeleccionar.AppendLine("	JOIN		Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE		Ordenes_Pagos.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("			AND	Retenciones_Documentos.Tip_Ori	= 'Ordenes_Pagos'")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.clase	= 'ISLR'")

            loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.factura BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		")

            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				''									AS Numero_Pago,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Control				    AS Control_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Factura				    AS Factura_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Fec_Ini				    AS Fecha_Factura,")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Net					AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Net					AS Monto_Abonado,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loComandoSeleccionar.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Cod_Pro				AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("				Proveedores.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Proveedores.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Proveedores.Dir_Fis					AS Direccion,")
            loComandoSeleccionar.AppendLine("               " & lcParametro4Desde & " AS Agrupar")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status IN  (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Documentos.factura BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro2Hasta)

            'loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rCRetencion_ISLRProveedores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCGS_rCRetencion_ISLRProveedores.ReportSource = loObjetoReporte


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
' RJG: 09/04/15: Codigo inicial, a partir de rCRetencion_ISLRProveedores.           		'
'-------------------------------------------------------------------------------------------'
