'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fReuniones_Eventos_Demos"
'-------------------------------------------------------------------------------------------'
Partial Class fReuniones_Eventos_Demos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Eventos_Marketing.Cod_Eve,")
            loConsulta.AppendLine("		    Eventos_Marketing.status,")
            loConsulta.AppendLine("		    Eventos_Marketing.Nom_Eve,")
            loConsulta.AppendLine("		    Eventos_Marketing.Responsable,")
            loConsulta.AppendLine("		    Eventos_Marketing.Fec_Ini,")
            loConsulta.AppendLine("		    Eventos_Marketing.Fec_Rev,")
            loConsulta.AppendLine("		    Eventos_Marketing.Fec_Fin,")
            loConsulta.AppendLine("		    Eventos_Marketing.Lugar,")
            loConsulta.AppendLine("		    Eventos_Marketing.Tipo,")
            loConsulta.AppendLine("		    Eventos_Marketing.Objetivo,")
            loConsulta.AppendLine("		    Eventos_Marketing.Por_Eje,")
            loConsulta.AppendLine("		    Eventos_Marketing.Duracion,")
            loConsulta.AppendLine("		    Eventos_Marketing.Nivel,")
            loConsulta.AppendLine("		    Eventos_Marketing.Etapa,")
            loConsulta.AppendLine("		    Eventos_Marketing.Clase,")
            loConsulta.AppendLine("		    Eventos_Marketing.prioridad,")
            loConsulta.AppendLine("		    Eventos_Marketing.Importancia,")
            loConsulta.AppendLine("		    Eventos_Marketing.Comentario")
            loConsulta.AppendLine("FROM		Eventos_Marketing")
            loConsulta.AppendLine("WHERE        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fReuniones_Eventos_Demos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfReuniones_Eventos_Demos.ReportSource = loObjetoReporte

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
