'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRetenciones_ISRLProveedores"
'-------------------------------------------------------------------------------------------'
Partial Class rRetenciones_ISRLProveedores
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

																				 
            loComandoSeleccionar.AppendLine("SELECT		Cuentas_Pagar.Fec_Ini		AS Fec_Ini,    	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Documento		AS Documento,  	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Cod_Pro		AS Cod_Pro,    	")
            loComandoSeleccionar.AppendLine("			Proveedores.Nom_Pro			AS Nom_Pro,    	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Tip_Ori		AS Tip_Ori,    	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Doc_Ori		AS Doc_Ori,    	")
            loComandoSeleccionar.AppendLine("			Cuentas_Pagar.Mon_Net		AS Mon_Net		")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Pagar")
            loComandoSeleccionar.AppendLine("	JOIN	Proveedores")
            loComandoSeleccionar.AppendLine("		ON	Proveedores.Cod_Pro = Cuentas_Pagar.Cod_Pro")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Pagar.Cod_tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND	" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Proveedores.Cod_Pro BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Pagar.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Pagar.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Pagar.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Proveedores.Cod_Zon BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Pagar.Cod_Rev BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro6Hasta)			 
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString(), "curReportes")

        '-------------------------------------------------------------------------------------------------------
        ' Verifica si el select (tabla nº0) trae registros
        '-------------------------------------------------------------------------------------------------------
            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If
			
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRetenciones_ISRLProveedores", laDatosReporte)
           
            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRetenciones_ISRLProveedores.ReportSource = loObjetoReporte


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
' CMS:  08/06/09: Codigo inicial															'
'-------------------------------------------------------------------------------------------'
' CMS:  14/07/09: Verificación de registros													'
'-------------------------------------------------------------------------------------------'
' RJG:  03/06/11: Eliminada la union con Vendedores en el SELECT, y ajustes en layout.		'
'-------------------------------------------------------------------------------------------'
' RJG:  11/09/14: Se ajustó el ancho de la columna "Número".                                '
'-------------------------------------------------------------------------------------------'
