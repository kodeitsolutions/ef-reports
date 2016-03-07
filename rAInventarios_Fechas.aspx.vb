'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rAInventarios_Fechas"
'-------------------------------------------------------------------------------------------'
Partial Class rAInventarios_Fechas
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1))
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro3Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("SELECT		Ajustes.Documento, ")
            loComandoSeleccionar.AppendLine("			Ajustes.Fec_Ini, ")
            loComandoSeleccionar.AppendLine("			CONVERT(NCHAR(30), Ajustes.Fec_Ini,112) AS Fec_Ini2, ")
            loComandoSeleccionar.AppendLine("			Ajustes.Comentario, ")
            loComandoSeleccionar.AppendLine("			Ajustes.Status, ")
            loComandoSeleccionar.AppendLine("			Ajustes.Cod_Mon, ")
            loComandoSeleccionar.AppendLine("			Ajustes.Tip_Ori, ")
            loComandoSeleccionar.AppendLine("			Ajustes.Doc_Ori, ")
            loComandoSeleccionar.AppendLine("			Ajustes.Mon_Net, ")
            loComandoSeleccionar.AppendLine("			Ajustes.Can_Art1 ")
			loComandoSeleccionar.AppendLine("FROM  		Ajustes ")
			loComandoSeleccionar.AppendLine("WHERE 		Ajustes.Fec_Ini			BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("			AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("			AND Ajustes.Documento	BETWEEN " & lcParametro1Desde)
			loComandoSeleccionar.AppendLine("			AND " & lcParametro1Hasta)
			loComandoSeleccionar.AppendLine("			AND Ajustes.Status		IN (" & lcParametro2Desde & ")")
			loComandoSeleccionar.AppendLine("			AND Ajustes.Tipo		IN (" & lcParametro3Desde & ")")
            loComandoSeleccionar.AppendLine("			AND Ajustes.Cod_Rev		BETWEEN " & lcParametro4Desde)
            loComandoSeleccionar.AppendLine("			AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("ORDER BY	Fec_Ini , " & lcOrdenamiento)


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

            
            
            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rAInventarios_Fechas", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)
            
            Me.crvrAInventarios_Fechas.ReportSource = loObjetoReporte
            

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
' JJD: 04/10/08: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
' GCR: 13/03/09: Estandarizacion de codigo y ajustes al diseño								'
'-------------------------------------------------------------------------------------------'
' CMS:  12/05/09: Ordenamiento																'
'-------------------------------------------------------------------------------------------'
' CMS:  14/05/09: Filtro “Revisión:”														'
'-------------------------------------------------------------------------------------------'
' CMS:  27/03/09: Filtro Tipo de Ajuste, se aplico el metodo de validacion de registro cero.'
'-------------------------------------------------------------------------------------------'
' RJG:  18/04/12: Se colocó el total de registros al fiunal del reporte.					'
'-------------------------------------------------------------------------------------------'
