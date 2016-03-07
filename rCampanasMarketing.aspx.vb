'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rCampanasMarketing"
'-------------------------------------------------------------------------------------------'
Partial Class rCampanasMarketing
    Inherits vis2Formularios.frmReporte

	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument  

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))

        Dim lcParametro1Desde AS String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde AS String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
        
		Try
            loConsulta.AppendLine("SELECT   Campanas.Documento AS Documento,")
            loConsulta.AppendLine("         Campanas.Nombre AS Nombre,")
            loConsulta.AppendLine("         Campanas.Responsable AS Responsable,")
            loConsulta.AppendLine("         Campanas.Fec_Ini,")
            loConsulta.AppendLine("         Campanas.Fec_Fin,")
            loConsulta.AppendLine("         Campanas.Por_Eje,")
            loConsulta.AppendLine("         Campanas.Status  AS Status")
            loConsulta.AppendLine("FROM	    Campanas")
            loConsulta.AppendLine("WHERE      Campanas.Documento BETWEEN " & lcParametro0Desde &" AND " & lcParametro0Hasta ) 
            loConsulta.AppendLine("     AND      Campanas.Etapa IN (" & lcParametro1Desde & " )")
            loConsulta.AppendLine("     AND      Campanas.Prioridad IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine("     AND      Campanas.Fec_Ini BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rCampanasMarketing", laDatosReporte)
			
			Me.mTraducirReporte(loObjetoReporte)
            
			Me.mFormatearCamposReporte(loObjetoReporte)

		    Me.crvrCampanasMarketing.ReportSource = loObjetoReporte
		    
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
' JAC : 28/07/15 : Codigo inicial
'-------------------------------------------------------------------------------------------'

