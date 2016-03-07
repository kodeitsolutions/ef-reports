'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCampaña"
'-------------------------------------------------------------------------------------------'
Partial Class fCampaña

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Campanas.Documento,")
            loConsulta.AppendLine("		    Campanas.status,")
            loConsulta.AppendLine("		    Campanas.Nombre,")
            loConsulta.AppendLine("		    Campanas.Responsable,")
            loConsulta.AppendLine("		    Campanas.Fec_Ini,")
            loConsulta.AppendLine("		    Campanas.Fec_Rev,")
            loConsulta.AppendLine("		    Campanas.Fec_Fin,")
            loConsulta.AppendLine("		    Campanas.Tipo,")
            loConsulta.AppendLine("		    Campanas.Objetivo,")
            loConsulta.AppendLine("		    Campanas.Por_Eje,")
            loConsulta.AppendLine("		    Campanas.Duracion,")
            loConsulta.AppendLine("		    Campanas.Nivel,")
            loConsulta.AppendLine("		    Campanas.Etapa,")
            loConsulta.AppendLine("		    Campanas.Clase,")
            loConsulta.AppendLine("		    Campanas.prioridad,")
            loConsulta.AppendLine("		    Campanas.Importancia,")
            loConsulta.AppendLine("		    Campanas.Comentario,")
            loConsulta.AppendLine("		    Campanas.Notas")
            loConsulta.AppendLine("FROM     Campanas")
            loConsulta.AppendLine("WHERE        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCampaña", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCampaña.ReportSource = loObjetoReporte

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
' RJG: 09/12/13: Programacion inicial
'-------------------------------------------------------------------------------------------'
