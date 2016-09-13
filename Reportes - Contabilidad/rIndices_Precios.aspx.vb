'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rIndices_Precios"
'-------------------------------------------------------------------------------------------'
Partial Class rIndices_Precios
     Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loComandoSeleccionar As New StringBuilder()
        
        Dim lnParametro0Desde As Integer = cusAplicacion.goReportes.paParametrosIniciales(0)
        Dim lnParametro0Hasta As Integer = cusAplicacion.goReportes.paParametrosFinales(0)
        
        If lnParametro0Desde <= 1900 Then 
			lnParametro0Desde = Date.Today.Year()
		End If
        If (lnParametro0Hasta <= 1900) OrElse (lnParametro0Hasta >= 99999999) Then 
			lnParametro0Hasta = Date.Today.Year()
		ElseIf(lnParametro0Hasta < lnParametro0Desde) Then 
			lnParametro0Hasta = lnParametro0Desde
		End If
		
		Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(lnParametro0Desde)
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(lnParametro0Hasta)
		
		Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
		
		Try	
		
			loComandoSeleccionar.AppendLine("SELECT		Año,")
			loComandoSeleccionar.AppendLine("			Por_Ene,")
			loComandoSeleccionar.AppendLine("			Mon_Ene,")
			loComandoSeleccionar.AppendLine("			Por_Feb,")
			loComandoSeleccionar.AppendLine("			Mon_Feb,")
			loComandoSeleccionar.AppendLine("			Por_Mar,")
			loComandoSeleccionar.AppendLine("			Mon_Mar,")
			loComandoSeleccionar.AppendLine("			Por_Abr,")
			loComandoSeleccionar.AppendLine("			Mon_Abr,")
			loComandoSeleccionar.AppendLine("			Por_May,")
			loComandoSeleccionar.AppendLine("			Mon_May,")
			loComandoSeleccionar.AppendLine("			Por_Jun,")
			loComandoSeleccionar.AppendLine("			Mon_Jun,")
			loComandoSeleccionar.AppendLine("			Por_Jul,")
			loComandoSeleccionar.AppendLine("			Mon_Jul,")
			loComandoSeleccionar.AppendLine("			Por_Ago,")
			loComandoSeleccionar.AppendLine("			Mon_Ago,")
			loComandoSeleccionar.AppendLine("			Por_Sep,")
			loComandoSeleccionar.AppendLine("			Mon_Sep,")
			loComandoSeleccionar.AppendLine("			Por_Oct,")
			loComandoSeleccionar.AppendLine("			Mon_Oct,")
			loComandoSeleccionar.AppendLine("			Por_Nov,")
			loComandoSeleccionar.AppendLine("			Mon_Nov,")
			loComandoSeleccionar.AppendLine("			Por_Dic,")
			loComandoSeleccionar.AppendLine("			Mon_Dic ")
			loComandoSeleccionar.AppendLine("FROM		Indices_Precios") 
			loComandoSeleccionar.AppendLine("WHERE		Año BETWEEN " & lcParametro0Desde)
			loComandoSeleccionar.AppendLine("		AND " & lcParametro0Hasta)
			loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
			loComandoSeleccionar.AppendLine("")
		
			'Me.mEscribirConsulta(loComandoSeleccionar.ToString())

		
            Dim loServicios As New cusDatos.goDatos
			
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If

			loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rIndices_Precios", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
			Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvrIndices_Precios.ReportSource = loObjetoReporte

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
' MJP: 15/07/08 : Codigo inicial															'
'-------------------------------------------------------------------------------------------' 
' MVP: 05/08/08: Cambios para multi idioma, mensaje de error y clase padre.					'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Estandarización de Código.													'
'-------------------------------------------------------------------------------------------'
' MAT: 13/05/11: Estandarización de Código.													'
'-------------------------------------------------------------------------------------------'
' RJG: 12/01/12: Ajuste en validacion de parámetro de Año. Ajuste visual.					'
'-------------------------------------------------------------------------------------------'
