'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCRetencion_IVAProveedores_CCMV_Multicentro"
'-------------------------------------------------------------------------------------------'
Partial Class rCRetencion_IVAProveedores_CCMV_Multicentro
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

            Dim lcAño As String = CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Year)
            Dim lcMes As String = Strings.Format(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Month, "00")
            'Dim lcMes As String = CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Month)
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT	    Cuentas_Pagar.Tip_Ori			                        	AS  Tipo_Origen,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Fec_Ini			                        	AS  Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Num_Com	                        	AS  Numero_Comprobante,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Documento			                        	AS  Numero_DocumentoRet,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Control			                        	AS  Numero_ControlRet,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Doc_Ori	                        	AS  Numero_Documento,")
            loComandoSeleccionar.AppendLine("			Documentos.Doc_Ori				                        	AS Doc_Pro001,")
            loComandoSeleccionar.AppendLine("			Documentos.Factura				                            AS Doc_Pro002,")
            'loComandoSeleccionar.AppendLine("			Documentos002.Factura				                        AS Doc_Pro002,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Doc_Ori				                        AS Doc_Pro002,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Cod_Tip	                        	AS  Tipo_Documento,")
            loComandoSeleccionar.AppendLine("			Documentos.Control				                        	AS  Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("			Renglones_Pagos.Mon_Net			                        	AS  Monto_Documento,")
            loComandoSeleccionar.AppendLine("			Documentos.Mon_Exe				                        	AS  Monto_ExentoDoc,")
            loComandoSeleccionar.AppendLine("			Documentos.Mon_Bas1				                        	AS  Monto_BaseDoc,")
            loComandoSeleccionar.AppendLine("			Documentos.Por_Imp1				                        	AS  Porcentaje_ImpuestoDoc,")
            loComandoSeleccionar.AppendLine("			Documentos.Mon_Imp1				                        	AS  Monto_ImpuestoDoc,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Mon_Bas	                        	AS  Base_Retencion,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Por_Ret	                        	AS  Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("			RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret     AS  Concepto,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Mon_Ret	                        	AS  Monto_Retenido,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Pro			                        	AS  Cod_Pro,")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro				                        	AS  Nom_Pro,")
            loComandoSeleccionar.AppendLine("			Proveedores.Rif					                        	AS  Rif,")
            loComandoSeleccionar.AppendLine("			Proveedores.Nit					                        	AS  Nit,")
            loComandoSeleccionar.AppendLine("			Proveedores.Dir_Fis				                        	AS  Direccion,")
            loComandoSeleccionar.AppendLine("			" & lcAño & "					                        	AS  Anio,")
            loComandoSeleccionar.AppendLine("           '" & lcMes & "'                                             AS  Mes")
            loComandoSeleccionar.AppendLine("FROM       Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("		    JOIN	    Pagos")
            loComandoSeleccionar.AppendLine("			ON          Pagos.Documento                     =   Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		    JOIN	    Retenciones_Documentos ")
            loComandoSeleccionar.AppendLine("			ON          Retenciones_Documentos.Documento    =   Pagos.Documento")
            loComandoSeleccionar.AppendLine("			AND         Retenciones_Documentos.Doc_Des      =   Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("			AND         Retenciones_Documentos.Cla_Des      =   Cuentas_Pagar.Cod_Tip")
            loComandoSeleccionar.AppendLine("			AND         Retenciones_Documentos.Tip_Des      =   'cuentas_pagar'")
            loComandoSeleccionar.AppendLine("			AND         Retenciones_Documentos.Clase	    =   'IMPUESTO'")
            loComandoSeleccionar.AppendLine("			AND         Retenciones_Documentos.Origen	    =   'Pagos'")
            loComandoSeleccionar.AppendLine("		    JOIN	    Renglones_Pagos ")
            loComandoSeleccionar.AppendLine("			ON          Renglones_Pagos.Documento           =   Pagos.Documento")
            loComandoSeleccionar.AppendLine("			AND         Renglones_Pagos.Doc_Ori             =   Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("	    	LEFT JOIN   Cuentas_Pagar AS Documentos")

            'loComandoSeleccionar.AppendLine("		    JOIN        Cuentas_Pagar                       AS  Documentos002 ")
            'loComandoSeleccionar.AppendLine("			ON          Documentos.Doc_Ori                =   Documentos002.Documento")

            loComandoSeleccionar.AppendLine("			ON          Documentos.Documento                =   Renglones_Pagos.Doc_Ori")
            loComandoSeleccionar.AppendLine("			AND         Documentos.Cod_Tip                  =   Renglones_Pagos.Cod_Tip")
            loComandoSeleccionar.AppendLine("	    	JOIN	    Proveedores")
            loComandoSeleccionar.AppendLine("			ON          Proveedores.Cod_Pro                 =   Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	    	LEFT JOIN   Retenciones")
            loComandoSeleccionar.AppendLine("			ON          Retenciones.Cod_Ret                 =   Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine(" WHERE		Cuentas_Pagar.Cod_Tip       =   'RETIVA'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status    <>  'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori   =   'Pagos'")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Fec_Ini   BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("        	AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Cod_Pro   BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Mon           BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Suc           BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro3Hasta)

            loComandoSeleccionar.AppendLine(" UNION ALL ")

            loComandoSeleccionar.AppendLine(" SELECT	Cuentas_Pagar.Tip_Ori				                        AS  Tipo_Origen,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Fec_Ini				                        AS  Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Num_Com	                        	AS  Numero_Comprobante,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Documento			                        	AS  Numero_DocumentoRet,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Control			                        	AS  Numero_ControlRet,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Doc_Ori	                        	AS  Numero_Documento,")
            loComandoSeleccionar.AppendLine("			Documentos.Doc_Ori				                        	AS Doc_Pro001,")
            loComandoSeleccionar.AppendLine("			Documentos.Factura			    	                        AS Doc_Pro002,")
            'loComandoSeleccionar.AppendLine("			Documentos002.Factura				                        AS Doc_Pro002,")
            'loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Doc_Ori				                        AS Doc_Pro002,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Cod_Tip	                        	AS  Tipo_Documento,")
            loComandoSeleccionar.AppendLine("			Documentos.Control				                        	AS  Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("			Documentos.Mon_Net				                        	AS  Monto_Documento,")
            loComandoSeleccionar.AppendLine("			Documentos.Mon_Exe				                        	AS  Monto_ExentoDoc,")
            loComandoSeleccionar.AppendLine("			Documentos.Mon_Bas1				                        	AS  Monto_BaseDoc,")
            loComandoSeleccionar.AppendLine("			Documentos.Por_Imp1				                        	AS  Porcentaje_ImpuestoDoc,")
            loComandoSeleccionar.AppendLine("			Documentos.Mon_Imp1				                        	AS  Monto_ImpuestoDoc,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Mon_Bas	                        	AS  Base_Retencion,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Por_Ret	                        	AS  Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("			RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret     AS  Concepto,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Mon_Ret	                        	AS  Monto_Retenido,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Pro			                        	AS  Cod_Pro,")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro				                        	AS  Nom_Pro,")
            loComandoSeleccionar.AppendLine("			Proveedores.Rif					                        	AS  Rif,")
            loComandoSeleccionar.AppendLine("			Proveedores.Nit					                        	AS  Nit,")
            loComandoSeleccionar.AppendLine("			Proveedores.Dir_Fis				                        	AS  Direccion,")
            loComandoSeleccionar.AppendLine("			" & lcAño & "					                        	AS  Anio,")
            loComandoSeleccionar.AppendLine("           '" & lcMes & "'                                             AS  Mes")
            loComandoSeleccionar.AppendLine(" FROM		Cuentas_Pagar")

            loComandoSeleccionar.AppendLine("		    JOIN        Cuentas_Pagar                       AS  Documentos ")
            loComandoSeleccionar.AppendLine("			ON          Documentos.Documento                =   Cuentas_Pagar.Doc_Ori")
            'loComandoSeleccionar.AppendLine("		    JOIN        Cuentas_Pagar                       AS  Documentos002 ")
            'loComandoSeleccionar.AppendLine("			ON          Documentos.Doc_Ori                =   Documentos002.Documento")

            loComandoSeleccionar.AppendLine("			AND         Documentos.Cod_Tip                  =   Cuentas_Pagar.Cla_Ori")
            loComandoSeleccionar.AppendLine("		    JOIN        Retenciones_Documentos")
            loComandoSeleccionar.AppendLine("			ON          Retenciones_Documentos.Doc_Des      =   Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("			AND         Retenciones_Documentos.Cla_Des      =   Cuentas_Pagar.Cod_Tip")
            loComandoSeleccionar.AppendLine("			AND         Retenciones_Documentos.Doc_Ori      =   Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("			AND         Retenciones_Documentos.Clase	    =   'IMPUESTO'")
            loComandoSeleccionar.AppendLine("	    	JOIN        Proveedores ")
            loComandoSeleccionar.AppendLine("			ON          Proveedores.Cod_Pro                 =   Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("		    LEFT JOIN   Retenciones")
            loComandoSeleccionar.AppendLine("			ON          Retenciones.Cod_Ret                 =   Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine(" WHERE		Cuentas_Pagar.Cod_Tip       =   'RETIVA'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status    <>  'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori   =   'Cuentas_Pagar'")
            loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Fec_Ini   BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       	AND Cuentas_Pagar.Cod_Pro   BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       	AND Cuentas_Pagar.Cod_Mon   BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       	AND Cuentas_Pagar.Cod_Suc   BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine(" ORDER BY " & lcOrdenamiento & ", Fecha_Retencion ASC")

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCRetencion_IVAProveedores_CCMV_Multicentro", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrCRetencion_IVAProveedores_CCMV_Multicentro.ReportSource = loObjetoReporte

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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 01/08/09: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 20/03/10: Agregado el filtro para que distinga retenciones de IVA de las de ISLR.	'
'-------------------------------------------------------------------------------------------'
' RJG: 01/06/11: Cambiado el número de control para que muestre el del documetno de origen	'
'				 en lugar del de la retención. Correcciones varias en la interface.		    '
'-------------------------------------------------------------------------------------------'
' JJD: 07/06/11: Ajustes a la generacion del comprobante. Formato del mes.                  '
'-------------------------------------------------------------------------------------------'
' JJD: 07/06/11: Busqueda de los valores de numero de factura y control.                    '
'-------------------------------------------------------------------------------------------'
' RJG: 04/04/14: Ajuste en la búsqueda del número de factura (registros duplicados).		'
'-------------------------------------------------------------------------------------------'
