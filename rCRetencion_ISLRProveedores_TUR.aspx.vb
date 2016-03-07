'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCRetencion_ISLRProveedores_TUR"
'-------------------------------------------------------------------------------------------'
Partial Class rCRetencion_ISLRProveedores_TUR
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

            Dim loConsulta As New StringBuilder()


            loConsulta.AppendLine("SELECT          Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loConsulta.AppendLine("                Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loConsulta.AppendLine("                Cuentas_Pagar.Doc_Ori				AS Numero_Pago,")
            loConsulta.AppendLine("                Facturas_Retenidas.Factura   	    AS Numero_Factura,")
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
            loConsulta.AppendLine("		JOIN	Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.clase = 'ISLR'")
            loConsulta.AppendLine("		JOIN	Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loConsulta.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine("        JOIN    Cuentas_Pagar Facturas_Retenidas")
            loConsulta.AppendLine("            ON  Facturas_Retenidas.Documento = Retenciones_Documentos.Doc_Ori")
            loConsulta.AppendLine("            AND Facturas_Retenidas.Cod_Tip = Retenciones_Documentos.Cod_Tip")
            loConsulta.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
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

            loConsulta.AppendLine("SELECT		Retenciones_Documentos.Tip_Ori		AS Tipo_Origen,")
            loConsulta.AppendLine("				Ordenes_Pagos.Fec_Ini				AS Fecha_Retencion,")
            loConsulta.AppendLine("				''									AS Numero_Pago,")
            loConsulta.AppendLine("             Ordenes_Pagos.Factura		        AS Numero_Factura,")
            loConsulta.AppendLine("             Ordenes_Pagos.Control			    AS Control_Factura,")
            loConsulta.AppendLine("				'ORD/PAG'							AS Tipo_Documento,")
            loConsulta.AppendLine("				Ordenes_Pagos.Documento				AS Numero_Documento,")
            loConsulta.AppendLine("				Ordenes_Pagos.Mon_Net				AS Monto_Documento,")
            loConsulta.AppendLine("				Ordenes_Pagos.Mon_Net				AS Monto_Abonado,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loConsulta.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loConsulta.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loConsulta.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loConsulta.AppendLine("				Ordenes_Pagos.Cod_Pro				AS Cod_Pro,")
            loConsulta.AppendLine("				Proveedores.Nom_Pro					AS Nom_Pro,")
            loConsulta.AppendLine("				Proveedores.Rif						AS Rif,")
            loConsulta.AppendLine("				Proveedores.Nit						AS Nit,")
            loConsulta.AppendLine("				Proveedores.Dir_Fis					AS Direccion")
            loConsulta.AppendLine("FROM			Retenciones_Documentos")
            loConsulta.AppendLine("	JOIN		Ordenes_Pagos ON Ordenes_Pagos.Documento = Retenciones_Documentos.documento")
            loConsulta.AppendLine("	JOIN		Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            loConsulta.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE		Ordenes_Pagos.Status = 'Confirmado'")
            loConsulta.AppendLine("			AND	Retenciones_Documentos.Tip_Ori	= 'Ordenes_Pagos'")
            loConsulta.AppendLine("			AND Retenciones_Documentos.clase	= 'ISLR'")

            loConsulta.AppendLine("           AND Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde)
            loConsulta.AppendLine("         		AND " & lcParametro0Hasta)
            loConsulta.AppendLine("           AND Ordenes_Pagos.Cod_Pro BETWEEN " & lcParametro1Desde)
            loConsulta.AppendLine("         		AND " & lcParametro1Hasta)
            loConsulta.AppendLine("           AND Ordenes_Pagos.Cod_Mon BETWEEN " & lcParametro2Desde)
            loConsulta.AppendLine("         		AND " & lcParametro2Hasta)
            loConsulta.AppendLine("           AND Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            loConsulta.AppendLine("         		AND " & lcParametro3Hasta)


            loConsulta.AppendLine("UNION ALL		")

            loConsulta.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loConsulta.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loConsulta.AppendLine("				''									AS Numero_Pago,")
            loConsulta.AppendLine("             Documentos.Factura		            AS Numero_Factura,")
            loConsulta.AppendLine("             Documentos.Control			        AS Control_Factura,")
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
            loConsulta.AppendLine("				Proveedores.Dir_Fis					AS Direccion")
            loConsulta.AppendLine("FROM			Cuentas_Pagar")
            loConsulta.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loConsulta.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loConsulta.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loConsulta.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loConsulta.AppendLine("	LEFT JOIN	Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loConsulta.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
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

            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCRetencion_ISLRProveedores_TUR", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCRetencion_ISLRProveedores_TUR.ReportSource = loObjetoReporte


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
' CMS: 21/05/09: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' CMS: 28/07/09: Se modificó la consulta de modo que se obtuvieron por separado los			'
'				 proveedores y los beneficiarios y luego se unieron los resultados.			'
'				 Verificacion de registros.													'
'				 Metodo de Ordenamiento														'
'-------------------------------------------------------------------------------------------'
' CMS: 29/07/09: Se Renonbre de Relación Global de ISLR Relativo a Relación Global de ISLR	'
'				 Retenido																	'
'-------------------------------------------------------------------------------------------'
' RJG: 20/03/10: Agregado el filtro para que distinga retenciones de IVA de las de ISLR.	'
'-------------------------------------------------------------------------------------------'
' JJD: 27/11/13: Se le agrego el Logo de la empresa                                         '
'-------------------------------------------------------------------------------------------'
' RJG: 27/02/15: Se le agrego El número de Factura (o documento) y controldel doc retenido. '
'-------------------------------------------------------------------------------------------'
