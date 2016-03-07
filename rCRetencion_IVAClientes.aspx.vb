'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCRetencion_IVAClientes"
'-------------------------------------------------------------------------------------------'
Partial Class rCRetencion_IVAClientes
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
			Dim lcMes As String = CStr(CDate(cusAplicacion.goReportes.paParametrosIniciales(0)).Month)
            Dim loComandoSeleccionar As New StringBuilder()




            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Cobrar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Num_Com		AS Numero_Comprobante,			")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Documento			AS Numero_DocumentoRet,			")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Control				AS Numero_ControlRet,			")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Control					AS Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("				Renglones_Cobros.Mon_Net			AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Exe					AS Monto_ExentoDoc,				")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Bas1					AS Monto_BaseDoc,				")
            loComandoSeleccionar.AppendLine("				Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,		")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Imp1					AS Monto_ImpuestoDoc,			")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loComandoSeleccionar.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Cod_Cli				AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli					AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Clientes.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Clientes.Dir_Fis					AS Direccion,")
            loComandoSeleccionar.AppendLine("				" & lcAño & "						AS Anio,")
            loComandoSeleccionar.AppendLine("				" & lcMes & "						AS Mes")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("		JOIN	Cobros")
            loComandoSeleccionar.AppendLine("			ON	Cobros.documento = Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ")
            loComandoSeleccionar.AppendLine("			ON	Retenciones_Documentos.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Cobrar.Documento")
            loComandoSeleccionar.AppendLine("		JOIN	Renglones_Cobros ")
            loComandoSeleccionar.AppendLine("			ON	Renglones_Cobros.Documento = Cobros.Documento")
            loComandoSeleccionar.AppendLine("			AND Renglones_Cobros.Doc_Ori = Retenciones_Documentos.Doc_Ori")
            loComandoSeleccionar.AppendLine("		LEFT JOIN	Cuentas_Cobrar AS Documentos										")
            loComandoSeleccionar.AppendLine("			ON	Documentos.Documento = Renglones_Cobros.Doc_Ori					")
            loComandoSeleccionar.AppendLine("			AND	Documentos.Cod_Tip = Renglones_Cobros.Cod_Tip 					")
            loComandoSeleccionar.AppendLine("		JOIN	Clientes")
            loComandoSeleccionar.AppendLine("			ON	Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("		LEFT JOIN Retenciones")
            loComandoSeleccionar.AppendLine("			ON	Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Cobrar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Tip_Ori = 'Cobros'")
											  
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("           AND Cuentas_Cobrar.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Mon BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("           AND Cobros.Cod_Suc BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL		")
            
            loComandoSeleccionar.AppendLine("SELECT			Cuentas_Cobrar.Tip_Ori				AS Tipo_Origen,")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Fec_Ini				AS Fecha_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Num_Com		AS Numero_Comprobante,				")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Documento			AS Numero_DocumentoRet,				")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Control				AS Numero_ControlRet,				")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Doc_Ori		AS Numero_Documento,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Cod_Tip		AS Tipo_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Control					AS Numero_ControlDoc,")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Net					AS Monto_Documento,")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Exe					AS Monto_ExentoDoc,					")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Bas1					AS Monto_BaseDoc,					")
            loComandoSeleccionar.AppendLine("				Documentos.Por_Imp1					AS Porcentaje_ImpuestoDoc,			")
            loComandoSeleccionar.AppendLine("				Documentos.Mon_Imp1					AS Monto_ImpuestoDoc,				")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Bas		AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Por_Ret		AS Porcentaje_Retenido,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Sus		AS Sustraendo_Retenido,")
            loComandoSeleccionar.AppendLine("				RTRIM(Retenciones.Cod_Ret) + ': ' + Retenciones.Nom_Ret	AS Concepto,")
            loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret		AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("				Cuentas_Cobrar.Cod_Cli				AS Cod_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Nom_Cli					AS Nom_Cli,")
            loComandoSeleccionar.AppendLine("				Clientes.Rif						AS Rif,")
            loComandoSeleccionar.AppendLine("				Clientes.Nit						AS Nit,")
            loComandoSeleccionar.AppendLine("				Clientes.Dir_Fis					AS Direccion,")
            loComandoSeleccionar.AppendLine("				" & lcAño & "						AS Anio,")
            loComandoSeleccionar.AppendLine("				" & lcMes & "						AS Mes")
            loComandoSeleccionar.AppendLine("FROM			Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("		JOIN Cuentas_Cobrar AS Documentos ")
            loComandoSeleccionar.AppendLine("			ON	Documentos.documento = Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Cobrar.Cla_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos")
            loComandoSeleccionar.AppendLine("			ON	Retenciones_Documentos.Doc_Des = Cuentas_Cobrar.Documento")
            loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Cobrar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		JOIN	Clientes ")
            loComandoSeleccionar.AppendLine("			ON	Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("		LEFT JOIN	Retenciones")
            loComandoSeleccionar.AppendLine("			ON Retenciones.Cod_Ret = Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE			Cuentas_Cobrar.Cod_Tip = 'RETIVA'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Status <> 'Anulado'")
            loComandoSeleccionar.AppendLine("			AND	Cuentas_Cobrar.Tip_Ori = 'Cuentas_Cobrar'")
											 
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Cobrar.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Cobrar.Cod_Mon BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       	    AND Cuentas_Cobrar.Cod_Suc BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("       	  		AND " & lcParametro3Hasta)

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCRetencion_IVAClientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrrCRetencion_IVAClientes.ReportSource = loObjetoReporte


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
' RJG: 01/08/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' RJG: 01/06/11: Cambiado el número de control para que muestre el del documetno de origen	'
'				 en lugar del de la retención. Correcciones varioas en la interface.		'
'-------------------------------------------------------------------------------------------'
