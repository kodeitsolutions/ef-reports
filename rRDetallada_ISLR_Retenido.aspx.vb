'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRDetallada_ISLR_Retenido"
'-------------------------------------------------------------------------------------------'
Partial Class rRDetallada_ISLR_Retenido
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




            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori					AS Tipo_Origen,							")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini					AS Fecha_Retencion,						")
            loComandoSeleccionar.AppendLine("				Detalles_Pagos.Tip_Ope					AS Tipo_Pago,							")
            loComandoSeleccionar.AppendLine("				CASE WHEN Detalles_Pagos.Tip_Ope='Efectivo'										")
            loComandoSeleccionar.AppendLine("					THEN 'Efectivo'																")
            loComandoSeleccionar.AppendLine("					ELSE Detalles_Pagos.Num_Doc													")
            loComandoSeleccionar.AppendLine("				END										AS Numero_Pago,							")
            loComandoSeleccionar.AppendLine("				Renglones_Pagos.Mon_Abo					AS Monto_Abonado,						")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas			AS Base_Retencion,						")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret			AS Porcentaje_Retenido,					")
            loComandoSeleccionar.AppendLine("				Retenciones.Cod_Ret						AS Codigo_Concepto,						")
            loComandoSeleccionar.AppendLine("				Retenciones.Nom_Ret						AS Concepto,							")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret			AS Monto_Retenido,						")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Cod_Pro					AS Cod_Pro,								")
            loComandoSeleccionar.AppendLine("				SUBSTRING(Proveedores.Nom_Pro, 0,30)	AS Nom_Pro,								")
            loComandoSeleccionar.AppendLine("				Proveedores.Rif							AS Rif,									")
            loComandoSeleccionar.AppendLine("				SUBSTRING(Proveedores.Dir_Fis, 0,25)	AS Direccion							")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("JOIN	Pagos ON Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("JOIN	Retenciones_Documentos ON Retenciones_Documentos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("				AND Retenciones_Documentos.doc_des = Cuentas_Pagar.documento")
            loComandoSeleccionar.AppendLine("JOIN	Renglones_Pagos ON Renglones_Pagos.Documento = Pagos.documento")
            loComandoSeleccionar.AppendLine("				AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("JOIN	Detalles_Pagos ON Detalles_Pagos.Documento = Pagos.Documento")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("				AND	Cuentas_Pagar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("				AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")
            loComandoSeleccionar.AppendLine("           	AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           	AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           	AND Pagos.Cod_Mon BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           	AND Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		")
            
            loComandoSeleccionar.AppendLine("SELECT			Retenciones_Documentos.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Fec_Ini						AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				Detalles_OPagos.Tip_Ope						AS Tipo_Pago,")
            loComandoSeleccionar.AppendLine("				CASE WHEN Detalles_OPagos.Tip_Ope='Efectivo'")
            loComandoSeleccionar.AppendLine("					THEN 'Efectivo'				")
            loComandoSeleccionar.AppendLine("					ELSE Detalles_OPagos.Num_Doc")
            loComandoSeleccionar.AppendLine("				END											AS Numero_Pago,			")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Mon_Net						AS Monto_Abonado,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas				AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret				AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				Retenciones.Cod_Ret							AS Codigo_Concepto,")
            loComandoSeleccionar.AppendLine("				Retenciones.Nom_Ret							AS Concepto,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret				AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Ordenes_Pagos.Cod_Pro						AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("				SUBSTRING(Proveedores.Nom_Pro, 0,30)		AS Nom_Pro,") 
            loComandoSeleccionar.AppendLine("				Proveedores.Rif								AS Rif,") 
            loComandoSeleccionar.AppendLine("				SUBSTRING(Proveedores.Dir_Fis, 0,25)	AS Direccion")
            loComandoSeleccionar.AppendLine("FROM		Retenciones_Documentos")
            loComandoSeleccionar.AppendLine("JOIN	Ordenes_Pagos ON Ordenes_Pagos.Documento = Retenciones_Documentos.documento")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores ON Proveedores.Cod_Pro = Ordenes_Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("LEFT JOIN Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("JOIN	Detalles_OPagos ON Detalles_OPagos.Documento = Ordenes_Pagos.Documento")
            loComandoSeleccionar.AppendLine("WHERE		Ordenes_Pagos.Status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("				AND	Retenciones_Documentos.Tip_Ori = 'Ordenes_Pagos'")
            loComandoSeleccionar.AppendLine("           	AND Ordenes_Pagos.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           	AND Ordenes_Pagos.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           	AND Ordenes_Pagos.Cod_Mon BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           	AND Ordenes_Pagos.Cod_Suc BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)
            

            loComandoSeleccionar.AppendLine("UNION ALL		")
            
            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Pagar.Tip_Ori					AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini					AS Fecha_Retencion,")
			loComandoSeleccionar.AppendLine("				'-'										AS Tipo_Pago,")
			loComandoSeleccionar.AppendLine("				'-'										AS Numero_Pago,")
			loComandoSeleccionar.AppendLine("				Documentos.Mon_Net						AS Monto_Abonado,")
			loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas			AS Base_Retencion,")
			loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret			AS Porcentaje_Retenido,")
			loComandoSeleccionar.AppendLine("				Retenciones.Cod_Ret						AS Codigo_Concepto,") 
			loComandoSeleccionar.AppendLine("				Retenciones.Nom_Ret						AS Concepto,") 
			loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret			AS Monto_Retenido,")
			loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Cod_Pro					AS Cod_Pro,")
			loComandoSeleccionar.AppendLine("				SUBSTRING(Proveedores.Nom_Pro, 0,30)	AS Nom_Pro,")
			loComandoSeleccionar.AppendLine("				Proveedores.Rif							AS Rif,")
			loComandoSeleccionar.AppendLine("				SUBSTRING(Proveedores.Dir_Fis, 0,25)	AS Direccion")
			loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
			loComandoSeleccionar.AppendLine("JOIN	Cuentas_Pagar AS Documentos ON Documentos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("				AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
            loComandoSeleccionar.AppendLine("JOIN	Retenciones_Documentos ON Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("				AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("JOIN	Proveedores ON Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("LEFT JOIN	Retenciones ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("				AND	Cuentas_Pagar.Status <> 'Anulado'") 
            loComandoSeleccionar.AppendLine("				AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'") 
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       	  	AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       	  	AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       	  	AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Pagar.Cod_Suc BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("       	  	AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY " & lcOrdenamiento)


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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRDetallada_ISLR_Retenido", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrrRDetallada_ISLRRProveedores.ReportSource = loObjetoReporte


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
