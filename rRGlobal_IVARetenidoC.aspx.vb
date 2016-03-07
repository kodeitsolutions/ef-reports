'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRGlobal_IVARetenidoC"
'-------------------------------------------------------------------------------------------'
Partial Class rRGlobal_IVARetenidoC
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
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)


            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


		'-------------------------------------------------------------------------------------------'
		' Retenciones de IVA generadas desde Pagos a Proveedores".									'
		'-------------------------------------------------------------------------------------------'
			loComandoSeleccionar.AppendLine("SELECT			Retenciones_Documentos.Mon_Bas			AS Base_Retencion,")
			loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret			AS Monto_Retenido,")
			loComandoSeleccionar.AppendLine("				Proveedores.Cod_Pro						AS Cod_Pro,")
			loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro						AS Nom_Pro,")
			'loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini					AS Fec_Ini,")
			loComandoSeleccionar.AppendLine("				Proveedores.Rif							AS Rif,")
			loComandoSeleccionar.AppendLine("				Proveedores.Nit							AS Nit,")
			loComandoSeleccionar.AppendLine("				CAST(Proveedores.Dir_Fis AS CHAR(200))	AS Direccion")
			loComandoSeleccionar.AppendLine("INTO			#tmpRetencionesIVA")
			loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
			loComandoSeleccionar.AppendLine("		LEFT JOIN Pagos")
			loComandoSeleccionar.AppendLine("			ON	Pagos.documento = Cuentas_Pagar.Doc_Ori")
			loComandoSeleccionar.AppendLine("			AND Pagos.status = 'Confirmado'")
			loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos ")
			loComandoSeleccionar.AppendLine("			ON	Retenciones_Documentos.Documento = Pagos.Documento")
			loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.doc_des = Cuentas_Pagar.Documento")
			loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Origen = 'Pagos'")
			loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Clase = 'IMPUESTO'")
			loComandoSeleccionar.AppendLine("		JOIN	Renglones_Pagos ")
			loComandoSeleccionar.AppendLine("			ON	Renglones_Pagos.Documento = Pagos.Documento")
			loComandoSeleccionar.AppendLine("			AND Renglones_Pagos.Doc_Ori = Retenciones_Documentos.Doc_Ori")
			loComandoSeleccionar.AppendLine("		LEFT JOIN Cuentas_Pagar AS Documentos								")
			loComandoSeleccionar.AppendLine("			ON	Documentos.Documento = Renglones_Pagos.Doc_Ori				")
			loComandoSeleccionar.AppendLine("			AND	Documentos.Cod_Tip = Renglones_Pagos.Cod_Tip 				")
			loComandoSeleccionar.AppendLine("		JOIN	Proveedores")
			loComandoSeleccionar.AppendLine("			ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
			loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'RETIVA'")
			loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
			loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'Pagos'")


			loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Pro BETWEEN " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Tip BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Cla BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Per BETWEEN " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine("         		AND " & lcParametro4Hasta)

			loComandoSeleccionar.AppendLine("UNION ALL		") 

               
 		'-------------------------------------------------------------------------------------------'
		' Retenciones de ISLR generadas desde Generar Documentos de Proveedores.					'
		'-------------------------------------------------------------------------------------------'
			loComandoSeleccionar.AppendLine("SELECT			Retenciones_Documentos.Mon_Bas			AS Base_Retencion,")
			loComandoSeleccionar.AppendLine("				Retenciones_Documentos.Mon_Ret			AS Monto_Retenido,")
			loComandoSeleccionar.AppendLine("				Proveedores.Cod_Pro						AS Cod_Pro,")
			loComandoSeleccionar.AppendLine("				Proveedores.Nom_Pro						AS Nom_Pro,")
			'loComandoSeleccionar.AppendLine("				Cuentas_Pagar.Fec_Ini					AS Fec_Ini,")
			loComandoSeleccionar.AppendLine("				Proveedores.Rif							AS Rif,")
			loComandoSeleccionar.AppendLine("				Proveedores.Nit							AS Nit,")
			loComandoSeleccionar.AppendLine("				CAST(Proveedores.Dir_Fis AS CHAR(200))	AS Direccion")
			loComandoSeleccionar.AppendLine("FROM			Cuentas_Pagar")
			loComandoSeleccionar.AppendLine("		LEFT JOIN Cuentas_Pagar AS Documentos ")
			loComandoSeleccionar.AppendLine("			ON	Documentos.documento = Cuentas_Pagar.Doc_Ori")
			loComandoSeleccionar.AppendLine("			AND Documentos.Cod_Tip = Cuentas_Pagar.Cla_Ori")
			loComandoSeleccionar.AppendLine("		JOIN	Retenciones_Documentos")
			loComandoSeleccionar.AppendLine("			ON	Retenciones_Documentos.Doc_Des = Cuentas_Pagar.Documento")
			loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Doc_Ori = Cuentas_Pagar.Doc_Ori")
			loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Tip_Ori		=	'Cuentas_Pagar'")
			loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Origen		=	'Cuentas_Pagar'")
			loComandoSeleccionar.AppendLine("			AND Retenciones_Documentos.Clase		=	'IMPUESTO'")
			loComandoSeleccionar.AppendLine("		JOIN	Proveedores ")
			loComandoSeleccionar.AppendLine("			ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
			loComandoSeleccionar.AppendLine("WHERE			Cuentas_Pagar.Cod_Tip = 'RETIVA'")
			loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Status <> 'Anulado'")
			loComandoSeleccionar.AppendLine("			AND	Cuentas_Pagar.Tip_Ori = 'cuentas_pagar'")

			loComandoSeleccionar.AppendLine("           AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("         		AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Pro BETWEEN " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("         		AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Tip BETWEEN " & lcParametro2Desde)
			loComandoSeleccionar.AppendLine("         		AND " & lcParametro2Hasta)
			loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Cla BETWEEN " & lcParametro3Desde)
			loComandoSeleccionar.AppendLine("         		AND " & lcParametro3Hasta)
			loComandoSeleccionar.AppendLine("           AND Proveedores.Cod_Per BETWEEN " & lcParametro4Desde)
			loComandoSeleccionar.AppendLine("         		AND " & lcParametro4Hasta)

			loComandoSeleccionar.AppendLine("SELECT		SUM(Base_Retencion) AS Mon_Bas,")
			loComandoSeleccionar.AppendLine("			SUM(Monto_Retenido) AS Mon_NetRet,")
			loComandoSeleccionar.AppendLine("			Cod_Pro,")
			loComandoSeleccionar.AppendLine("			Nom_Pro,")
			loComandoSeleccionar.AppendLine("			Rif					AS Rif_Pro,")
			loComandoSeleccionar.AppendLine("			Nit,")
			loComandoSeleccionar.AppendLine("			Direccion")
			loComandoSeleccionar.AppendLine("FROM		#tmpRetencionesIVA")
			loComandoSeleccionar.AppendLine("GROUP BY	Cod_Pro, Nom_Pro, Rif, Nit, Direccion")
            loComandoSeleccionar.AppendLine("ORDER BY  " & lcOrdenamiento)



            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

            '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRGlobal_IVARetenidoC", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRGlobal_IVARetenidoC.ReportSource = loObjetoReporte


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
' RJG: 17/03/11: Ajuste en los filtros de las retenciones
'-------------------------------------------------------------------------------------------'
