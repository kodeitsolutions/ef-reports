'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fReposos"
'-------------------------------------------------------------------------------------------'
Partial Class fReposos

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Reposos.Documento,")
            loConsulta.AppendLine("		    Reposos.registro AS Fecha,")
            loConsulta.AppendLine("		    Reposos.status,")
            loConsulta.AppendLine("		    Reposos.cod_tra,")
            loConsulta.AppendLine("	    	Trabajadores.nom_tra,")
            loConsulta.AppendLine("	    	Trabajadores.cedula,")
            loConsulta.AppendLine("	    	Reposos.remunerado,")
            loConsulta.AppendLine("		    Reposos.Fec_ini AS Desde,")
            loConsulta.AppendLine("	    	Reposos.Fec_Fin AS hasta,")
            loConsulta.AppendLine("		    Reposos.Dias,")
            loConsulta.AppendLine("	    	Reposos.justificado,")
            loConsulta.AppendLine("	    	Reposos.comentario,")
            loConsulta.AppendLine("	    	Reposos.motivo,")
            loConsulta.AppendLine("	    	Motivos.nom_mot")
            loConsulta.AppendLine("FROM     Reposos")
            loConsulta.AppendLine(" JOIN    Trabajadores ON Trabajadores.cod_Tra = Reposos.Cod_Tra")
            loConsulta.AppendLine(" JOIN    Motivos ON Motivos.cod_mot = Reposos.motivo")
            loConsulta.AppendLine("WHERE        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fReposos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfReposos.ReportSource = loObjetoReporte

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
