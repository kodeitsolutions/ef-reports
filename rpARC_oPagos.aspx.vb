'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rpARC_oPagos"
'-------------------------------------------------------------------------------------------'
Partial Class rpARC_oPagos
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
		' Retenciones de ISLR generadas desde Pagos a Proveedores".									'
		'-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine("SELECT		Retenciones_Documentos.Mon_Bas			AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Mon_Ret			AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Tip_Ori			AS Tip_Ori,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Doc_Ori			AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Cla_Ori			AS Tip_Doc,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Fec_Ini					AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("			Proveedores.Cod_Pro						AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro						AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("			Proveedores.Rif							AS Rif,") 
            loComandoSeleccionar.AppendLine("			Proveedores.Nit							AS Nit,") 
            loComandoSeleccionar.AppendLine("			CAST(Proveedores.Dir_Fis AS CHAR(200))	AS Direccion") 
            loComandoSeleccionar.AppendLine("INTO		#tmpRetencionesISLR") 
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("	JOIN	Pagos") 
            loComandoSeleccionar.AppendLine("		ON  Pagos.documento = Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		AND Pagos.status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("	JOIN	Retenciones_Documentos") 
            loComandoSeleccionar.AppendLine("		ON  Retenciones_Documentos.Documento	= Pagos.documento")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.Doc_Des		= Cuentas_Pagar.documento")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.Origen = 'Pagos'")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.Clase = 'ISLR'")
			loComandoSeleccionar.AppendLine("    JOIN	Renglones_Pagos") 
			loComandoSeleccionar.AppendLine("		ON  Renglones_Pagos.Documento			= Pagos.documento")
			loComandoSeleccionar.AppendLine("		AND Renglones_Pagos.Doc_Ori				= Retenciones_Documentos.Doc_Ori")
			loComandoSeleccionar.AppendLine("	JOIN	Proveedores") 
            loComandoSeleccionar.AppendLine("		ON  Proveedores.Cod_Pro             	= Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Retenciones") 
            loComandoSeleccionar.AppendLine("		ON  Retenciones.Cod_Ret             	= Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Pagar.Cod_Tip           	= 'ISLR'")
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Status				<> 'Anulado'")
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Tip_Ori				= 'Pagos' ")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Fec_Ini       		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Pro         		BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Tip         		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Cla         		BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("		AND Proveedores.Cod_Per         		BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro4Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL	") 

 		'-------------------------------------------------------------------------------------------'
		' Retenciones de ISLR generadas desde Ordenes de Pago".										'
		'-------------------------------------------------------------------------------------------'
           loComandoSeleccionar.AppendLine("SELECT		Retenciones_Documentos.Mon_Bas			AS Base_Retencion,") 
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Mon_Ret			AS Monto_Retenido,") 
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Tip_Ori			AS Tip_Ori,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Documento		AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Cla_Ori			AS Tip_Doc,")
            loComandoSeleccionar.AppendLine("			Ordenes_Pagos.Fec_Ini                   AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("			Proveedores.Cod_Pro						AS Cod_Pro,") 
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro						AS Nom_Pro,") 
            loComandoSeleccionar.AppendLine("			Proveedores.Rif							AS Rif,") 
            loComandoSeleccionar.AppendLine("			Proveedores.Nit							AS Nit,") 
            loComandoSeleccionar.AppendLine("			CAST(Proveedores.Dir_Fis AS CHAR(200))	AS Direccion")
            loComandoSeleccionar.AppendLine("FROM		Retenciones_Documentos")
            loComandoSeleccionar.AppendLine("	JOIN	Ordenes_Pagos") 
            loComandoSeleccionar.AppendLine("		ON	Ordenes_Pagos.Documento             =   Retenciones_Documentos.Documento")
            loComandoSeleccionar.AppendLine("		AND	Ordenes_Pagos.Status				=   'Confirmado'")
            loComandoSeleccionar.AppendLine("	JOIN	Proveedores") 
            loComandoSeleccionar.AppendLine("		ON	Proveedores.Cod_Pro                 =   Ordenes_Pagos.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN Retenciones") 
            loComandoSeleccionar.AppendLine("		ON Retenciones.Cod_Ret                  =   Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE		Retenciones_Documentos.Tip_Ori		= 'Ordenes_Pagos'")
            loComandoSeleccionar.AppendLine("		AND	Retenciones_Documentos.Clase		= 'ISLR'")
            loComandoSeleccionar.AppendLine("     	AND Ordenes_Pagos.Fec_Ini           	BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("     		AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("     	AND Proveedores.Cod_Pro             	BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("     		AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("     	AND Proveedores.Cod_Tip             	BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("     		AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("     	AND Proveedores.Cod_Cla             	BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("     		AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("     	AND Proveedores.Cod_Per             	BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("     		AND " & lcParametro4Hasta)

            loComandoSeleccionar.AppendLine("UNION ALL ")
            
 		'-------------------------------------------------------------------------------------------'
		' Retenciones de ISLR generadas desde Generar Documentos de Proveedores".					'
		'-------------------------------------------------------------------------------------------'
            loComandoSeleccionar.AppendLine("SELECT		Retenciones_Documentos.Mon_Bas			AS Base_Retencion,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Mon_Ret			AS Monto_Retenido,")
            loComandoSeleccionar.AppendLine("           Retenciones_Documentos.Tip_Ori			AS Tip_Ori,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Doc_Ori			AS Doc_Ori,")
            loComandoSeleccionar.AppendLine("			Retenciones_Documentos.Cla_Ori			AS Tip_Doc,")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Fec_Ini                   AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("			Proveedores.Cod_Pro						AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro						AS Nom_Pro,")
            loComandoSeleccionar.AppendLine("			Proveedores.Rif							AS Rif,")
            loComandoSeleccionar.AppendLine("			Proveedores.Nit							AS Nit,")
            loComandoSeleccionar.AppendLine("			CAST(Proveedores.Dir_Fis AS CHAR(200))	AS Direccion")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("	JOIN	Retenciones_Documentos") 
            loComandoSeleccionar.AppendLine("		ON  Retenciones_Documentos.Doc_Des  	=   Cuentas_Pagar.Documento")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.Doc_Ori  	=   Cuentas_Pagar.Doc_Ori")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.Tip_Ori		=	'Cuentas_Pagar'")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.Origen		=	'Cuentas_Pagar'")
            loComandoSeleccionar.AppendLine("		AND Retenciones_Documentos.Clase		=	'ISLR'")
            loComandoSeleccionar.AppendLine("	JOIN    Proveedores") 
            loComandoSeleccionar.AppendLine("		ON  Proveedores.Cod_Pro             	=   Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("	LEFT JOIN   Retenciones") 
            loComandoSeleccionar.AppendLine("		ON  Retenciones.Cod_Ret             	=   Retenciones_Documentos.Cod_Ret")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Pagar.Cod_Tip           	=   'ISLR'")
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Status        		<>  'Anulado'")
            loComandoSeleccionar.AppendLine("		AND	Cuentas_Pagar.Tip_Ori       		=   'Cuentas_Pagar'")
            loComandoSeleccionar.AppendLine("       AND Cuentas_Pagar.Fec_Ini       		BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Proveedores.Cod_Pro         		BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Proveedores.Cod_Tip         		BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Proveedores.Cod_Cla         		BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("       AND Proveedores.Cod_Per         		BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("         	AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY	Proveedores.Cod_Pro")



            loComandoSeleccionar.AppendLine("SELECT		Cuentas_Pagar.Mon_Net                   AS  Mon_Net,")
            loComandoSeleccionar.AppendLine("       	(#tmpRetencionesISLR.Base_Retencion)    AS  Mon_Bas,")
            loComandoSeleccionar.AppendLine("			(#tmpRetencionesISLR.Monto_Retenido)    AS  Mon_NetRet,")
            loComandoSeleccionar.AppendLine("			'Compras'                               AS  Tip_Ori,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Doc_Ori             AS  Doc_Ori,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Tip_Doc             AS  Tip_Doc,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Cod_Pro             AS  Cod_Pro,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Nom_Pro             AS  Nom_Pro,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Fec_Ini             AS  Fec_Ini,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Rif                 AS  Rif_Pro,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Nit                 AS  Nif_Pro,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Direccion           AS  Direccion")
            loComandoSeleccionar.AppendLine("INTO   	#tmpRetencionesISLR002")
            loComandoSeleccionar.AppendLine("FROM		#tmpRetencionesISLR, Cuentas_Pagar ")
            loComandoSeleccionar.AppendLine("WHERE  	#tmpRetencionesISLR.Tip_Doc     =   Cuentas_Pagar.Cod_Tip ")
            loComandoSeleccionar.AppendLine("   	AND #tmpRetencionesISLR.Doc_Ori		=   Cuentas_Pagar.Documento ")
            loComandoSeleccionar.AppendLine("   	AND #tmpRetencionesISLR.Tip_Ori		=   'Cuentas_Pagar' ")

            loComandoSeleccionar.AppendLine("UNION ALL ")

            loComandoSeleccionar.AppendLine("SELECT		Ordenes_Pagos.Mon_Net                   AS  Mon_Net,")
            loComandoSeleccionar.AppendLine("       	(#tmpRetencionesISLR.Base_Retencion)    AS  Mon_Bas,")
            loComandoSeleccionar.AppendLine("			(#tmpRetencionesISLR.Monto_Retenido)    AS  Mon_NetRet,")
            loComandoSeleccionar.AppendLine("			'Ordenes'                               AS  Tip_Ori,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Doc_Ori             AS  Doc_Ori,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Tip_Ori             AS  Tip_Doc,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Cod_Pro             AS  Cod_Pro,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Nom_Pro             AS  Nom_Pro,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Fec_Ini             AS  Fec_Ini,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Rif                 AS  Rif_Pro,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Nit                 AS  Nif_Pro,")
            loComandoSeleccionar.AppendLine("			#tmpRetencionesISLR.Direccion           AS  Direccion")
            loComandoSeleccionar.AppendLine("FROM		#tmpRetencionesISLR, Ordenes_Pagos")
            loComandoSeleccionar.AppendLine("WHERE 		#tmpRetencionesISLR.Doc_Ori     =   Ordenes_Pagos.Documento ")
            loComandoSeleccionar.AppendLine("      		AND #tmpRetencionesISLR.Tip_Ori =  'Ordenes_Pagos' ")
            loComandoSeleccionar.AppendLine("ORDER BY  " & lcOrdenamiento & ",#tmpRetencionesISLR.Fec_Ini")

            loComandoSeleccionar.AppendLine("SELECT		Mon_Net,")
            loComandoSeleccionar.AppendLine("           Mon_Bas,")
            loComandoSeleccionar.AppendLine("			Mon_NetRet,")
            loComandoSeleccionar.AppendLine("			Tip_Ori,")
            loComandoSeleccionar.AppendLine("			Doc_Ori,")
            loComandoSeleccionar.AppendLine("			Tip_Doc,")
            loComandoSeleccionar.AppendLine("			Cod_Pro,")
            loComandoSeleccionar.AppendLine("			Nom_Pro,")
            loComandoSeleccionar.AppendLine("			Fec_Ini,")
            loComandoSeleccionar.AppendLine("			Rif_Pro,")
            loComandoSeleccionar.AppendLine("			Nif_Pro,")
            loComandoSeleccionar.AppendLine("			Direccion")
            loComandoSeleccionar.AppendLine("FROM		#tmpRetencionesISLR002")

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rpARC_oPagos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrpARC_oPagos.ReportSource = loObjetoReporte


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
' CMS: 21/05/09: Codigo inicial
'-------------------------------------------------------------------------------------------'
' CMS: 28/07/09: Se modofico la consulta de modo que se obtuvieron por separado los
'                proveedores y los beneficiarios y luego se unieron los resultados.
'                Verificacion de registros.
'                Metodo de Ordenamiento
'-------------------------------------------------------------------------------------------'
' CMS: 29/07/09: Se Renonbre de Relación Global de ISLR Relativo a Relación Global de ISLR 
'                Retenido
'-------------------------------------------------------------------------------------------'
' JJD: 30/03/10: Ajustes a los campos obtenidos
'-------------------------------------------------------------------------------------------'
' JJD: 31/03/10: Ajustes a los campos obtenidos
'-------------------------------------------------------------------------------------------'
' RJG: 17/03/11: Ajuste en los filtros de las retenciones
'-------------------------------------------------------------------------------------------'
