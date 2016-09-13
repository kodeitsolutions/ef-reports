'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fInstancia_Global"
'-------------------------------------------------------------------------------------------'
Partial Class fInstancia_Global

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Factory_Global.dbo.Clientes.Cod_Cli,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.Nom_Cli,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.status,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.rif,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.Fec_Ini,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.Fec_Fin,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.Contacto,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.Correo,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.web,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.telefonos,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.movil,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.control,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.dir_fis,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.Referencia,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.Cod_Emp,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.Cod_Usu,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Clientes.Cod_Gru")
            loConsulta.AppendLine("FROM     Factory_Global.dbo.Clientes ")
            loConsulta.AppendLine("WHERE        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fInstancia_Global", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfInstancia_Global.ReportSource = loObjetoReporte

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
' EAG: 25/08/15: Programacion inicial
'-------------------------------------------------------------------------------------------'
