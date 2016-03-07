Imports System.Data
Partial Class rPagos_FechasResumido

    Inherits vis2Formularios.frmReporte
    
    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

	Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

		Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = cusAplicacion.goReportes.paParametrosIniciales(7).ToString().Trim()

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("SELECT		Pagos.Documento						AS Documento,")
            loComandoSeleccionar.AppendLine("			Pagos.Status						AS Status,")
            loComandoSeleccionar.AppendLine("			Pagos.Fec_Ini						AS Fec_Ini,")
            loComandoSeleccionar.AppendLine("			Pagos.Cod_Pro						AS Cod_Pro,")
            loComandoSeleccionar.AppendLine("			SUBSTRING(Proveedores.Nom_Pro,1,40) AS Nom_Pro, ")
            loComandoSeleccionar.AppendLine("			Pagos.Comentario					AS Comentario,")
            loComandoSeleccionar.AppendLine("			Pagos.Mon_Net						AS Neto, ")
            loComandoSeleccionar.AppendLine("			Pagos.Cod_Mon						AS Cod_Mon, ")
            loComandoSeleccionar.AppendLine("			Pagos.Mon_Sal						AS Mon_Sal, ")
            loComandoSeleccionar.AppendLine("			CONVERT(VARCHAR(20), Pagos.Fec_Ini, 112) AS Fecha_Orden ")
            
            loComandoSeleccionar.AppendLine("FROM		Pagos")
            loComandoSeleccionar.AppendLine("	JOIN	Proveedores ON Pagos.Cod_Pro = Proveedores.Cod_Pro ")
            loComandoSeleccionar.AppendLine("WHERE			Pagos.Documento     BETWEEN " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Fec_Ini       BETWEEN " & lcParametro1Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Cod_Pro       BETWEEN " & lcParametro2Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro2Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Cod_Mon       BETWEEN " & lcParametro3Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("			AND Pagos.Status        IN (" & lcParametro4Desde & ")")
            
            loComandoSeleccionar.AppendLine("           AND Pagos.Cod_Suc    BETWEEN " & lcParametro5Desde)
            loComandoSeleccionar.AppendLine("           AND " & lcParametro5Hasta)
            
            If lcParametro7Desde = "Igual" Then
                loComandoSeleccionar.AppendLine("           AND Pagos.Cod_rev   BETWEEN " & lcParametro6Desde)
            Else
                loComandoSeleccionar.AppendLine("           AND Pagos.Cod_rev  NOT BETWEEN " & lcParametro6Desde)
            End If

            loComandoSeleccionar.AppendLine("           AND " & lcParametro6Hasta)
																												   
            loComandoSeleccionar.AppendLine("ORDER BY    Fecha_Orden, " & lcOrdenamiento)

			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

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

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPagos_FechasResumido", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPagos_FechasResumido.ReportSource = loObjetoReporte

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
' RJG: 10/04/12: Programacion inicial.														'
'-------------------------------------------------------------------------------------------'
