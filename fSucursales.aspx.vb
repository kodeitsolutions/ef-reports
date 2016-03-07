'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fSucursales"
'-------------------------------------------------------------------------------------------'
Partial Class fSucursales

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	sucursales.cod_suc,")
            loConsulta.AppendLine("         (CASE   sucursales.status")
            loConsulta.AppendLine("		            WHEN 'A' THEN 'ACTIVO'")
            loConsulta.AppendLine("		            WHEN 'I' THEN 'INACTIVO'")
            loConsulta.AppendLine("		            WHEN 'S' THEN 'SUSPENDIDO'")
            loConsulta.AppendLine("		    END) AS status,")
            loConsulta.AppendLine("		sucursales.nom_suc,")
            loConsulta.AppendLine("		sucursales.contacto,")
            loConsulta.AppendLine("		sucursales.nivel,")
            loConsulta.AppendLine("		sucursales.correo,")
            loConsulta.AppendLine("		sucursales.telefonos,")
            loConsulta.AppendLine("		sucursales.Movil,")
            loConsulta.AppendLine("		sucursales.direccion,")
            loConsulta.AppendLine("		sucursales.con_ini AS Desde,")
            loConsulta.AppendLine("		sucursales.con_fin AS Hasta,")
            loConsulta.AppendLine("		sucursales.fec_ini,")
            loConsulta.AppendLine("		sucursales.fec_fin,")
            loConsulta.AppendLine("		sucursales.comentario")
            loConsulta.AppendLine("FROM Sucursales")
            loConsulta.AppendLine("WHERE        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fSucursales", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfSucursales.ReportSource = loObjetoReporte

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
