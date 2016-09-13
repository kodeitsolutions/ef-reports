'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPermisos"
'-------------------------------------------------------------------------------------------'
Partial Class fPermisos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Permisos.Documento,")
            loConsulta.AppendLine("		    Permisos.registro AS Fecha,")
            loConsulta.AppendLine("		    Permisos.status,")
            loConsulta.AppendLine("		    Permisos.cod_tra,")
            loConsulta.AppendLine("	    	Trabajadores.nom_tra,")
            loConsulta.AppendLine("	    	Trabajadores.cedula,")
            loConsulta.AppendLine("	    	Permisos.remunerado,")
            loConsulta.AppendLine("		    Permisos.Fec_ini AS Desde,")
            loConsulta.AppendLine("	    	Permisos.Fec_Fin AS hasta,")
            loConsulta.AppendLine("		    Permisos.Dias,")
            loConsulta.AppendLine("	    	Permisos.justificado,")
            loConsulta.AppendLine("	    	Permisos.comentario,")
            loConsulta.AppendLine("	    	Permisos.motivo,")
            loConsulta.AppendLine("	    	Motivos.nom_mot")
            loConsulta.AppendLine("FROM     Permisos")
            loConsulta.AppendLine(" JOIN    Trabajadores ON Trabajadores.cod_Tra = Permisos.Cod_Tra")
            loConsulta.AppendLine(" JOIN    Motivos ON Motivos.cod_mot = Permisos.motivo")
            loConsulta.AppendLine("WHERE        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPermisos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfPermisos.ReportSource = loObjetoReporte

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
