'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rRetenciones_ISRLClientes"
'-------------------------------------------------------------------------------------------'
Partial Class rRetenciones_ISRLClientes
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


            loComandoSeleccionar.AppendLine("SELECT		Cuentas_Cobrar.Fec_Ini		AS Fec_Ini,    	")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Documento	AS Documento,  	")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Cod_Cli		AS Cod_Cli,    	")
            loComandoSeleccionar.AppendLine("			Clientes.Nom_Cli			AS Nom_Cli,    	")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Tip_Ori		AS Tip_Ori,    	")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Doc_Ori		AS Doc_Ori,    	")
            loComandoSeleccionar.AppendLine("			Cuentas_Cobrar.Mon_Net		AS Mon_Net		")
            loComandoSeleccionar.AppendLine("FROM		Cuentas_Cobrar")
            loComandoSeleccionar.AppendLine("	JOIN	Clientes")
            loComandoSeleccionar.AppendLine("		ON	Clientes.Cod_Cli = Cuentas_Cobrar.Cod_Cli")
            loComandoSeleccionar.AppendLine("WHERE		Cuentas_Cobrar.Cod_tip = 'ISLR'")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Fec_Ini BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND	" & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("       AND Clientes.Cod_Cli BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Ven BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Status IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("		AND Cuentas_Cobrar.Cod_Mon BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("       AND Clientes.Cod_Zon BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("       AND Cuentas_Cobrar.Cod_Rev BETWEEN " & lcParametro6Desde)
            loComandoSeleccionar.AppendLine("       	AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY   " & lcOrdenamiento)		  


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodos(loComandoSeleccionar.ToString, "curReportes")

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
			
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rRetenciones_ISRLClientes", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrRetenciones_ISRLClientes.ReportSource = loObjetoReporte


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
' RJG:  03/06/11: Eliminada la union con Vendedores en el SELECT, y ajustes en layout.		'
'-------------------------------------------------------------------------------------------'
