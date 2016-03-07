'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCRetencion_ISLRProveedores_CCSJ_002"
'-------------------------------------------------------------------------------------------'
Partial Class rCRetencion_ISLRProveedores_CCSJ_002
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

            Dim loComandoSeleccionar As New StringBuilder()




            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Doc_Ori				AS Numero_Pago,")
            loComandoSeleccionar.AppendLine("				Origenes.Factura					AS Numero_Proveedor,")
            loComandoSeleccionar.AppendLine("				Origenes.Control					AS Control_Proveedor,")
            loComandoSeleccionar.AppendLine("				Origenes.Fec_Ini                    AS Fecha_Proveedor,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
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
            loComandoSeleccionar.AppendLine("				Pagos.Recibo                        AS Recibo_Documento,")
            loComandoSeleccionar.AppendLine("				Pagos.Control                       AS Control_Documento")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("		JOIN	Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.clase = 'ISLR'")
            loComandoSeleccionar.AppendLine("		JOIN	Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("		JOIN    Cuentas_Pagar AS Origenes ON Origenes.Cod_Tip = Retenciones_Documentos.Cod_Tip")
            loComandoSeleccionar.AppendLine("			AND Origenes.Documento = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")

            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Mon BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		")

            loComandoSeleccionar.AppendLine("SELECT			Retenciones_Documentos.Tip_Ori		AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				''									AS Numero_Pago,")
            loComandoSeleccionar.AppendLine("				''									AS Numero_Proveedor,")
            loComandoSeleccionar.AppendLine("				''									AS Control_Proveedor,")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Fec_Ini				AS Fecha_Proveedor,")
            loComandoSeleccionar.AppendLine("				'ORD/PAG'							AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Documento				AS Numero_Documento,")
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
            loComandoSeleccionar.AppendLine("				''                                  AS Recibo_Documento,")
            loComandoSeleccionar.AppendLine("				''                                  AS Control_Documento")
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
            loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Cod_Mon BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)


            loComandoSeleccionar.AppendLine("UNION ALL		")

            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				''									AS Numero_Pago,")
            loComandoSeleccionar.AppendLine("				''									AS Numero_Proveedor,")
            loComandoSeleccionar.AppendLine("				''									AS Control_Proveedor,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini				AS Fecha_Proveedor,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
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
            loComandoSeleccionar.AppendLine("				''                                  AS Recibo_Documento,")
            loComandoSeleccionar.AppendLine("				''                                  AS Control_Documento")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("		JOIN	Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN	Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro3Hasta)

            loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCRetencion_ISLRProveedores_CCSJ_002", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCRetencion_ISLRProveedores_CCSJ_002.ReportSource = loObjetoReporte


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
' JJD: 18/10/13: Se le agrego los campos de Num de Factura, de Control y Fecha de Factura   '
'-------------------------------------------------------------------------------------------'
' JJD: 21/11/13: Programacion inicial del formato de Retenciones de Proveedores para CCSJ   '
'-------------------------------------------------------------------------------------------'
