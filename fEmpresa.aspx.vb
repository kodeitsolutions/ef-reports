'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fEmpresa"
'-------------------------------------------------------------------------------------------'
Partial Class fEmpresa

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Factory_Global.dbo.empresas.cod_emp,")
            loConsulta.AppendLine("         (CASE   Factory_Global.dbo.empresas.status")
            loConsulta.AppendLine("		            WHEN 'A' THEN 'ACTIVO'")
            loConsulta.AppendLine("		            WHEN 'I' THEN 'INACTIVO'")
            loConsulta.AppendLine("		            WHEN 'S' THEN 'SUSPENDIDO'")
            loConsulta.AppendLine("		    END) AS status,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.nom_emp,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.mul_suc,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.tip_rep,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.rif,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.nit,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.tipo,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.nivel,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.sistema,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.cod_neg,")
            loConsulta.AppendLine("		COALESCE(Factory_Global.dbo.Negocios.nom_neg,'') AS nom_reg,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.fec_ini,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.fec_fin,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.contacto,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.correo,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.web,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.telefonos,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.fax,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.movil,")
            loConsulta.AppendLine("		Factory_Global.dbo.empresas.direccion,")
            loConsulta.AppendLine("        Factory_Global.dbo.empresas.comentario")
            loConsulta.AppendLine("FROM Factory_Global.dbo.empresas")
            loConsulta.AppendLine("LEFT JOIN Factory_Global.dbo.Negocios ON Factory_Global.dbo.negocios.cod_neg= Factory_Global.dbo.empresas.cod_neg")
            loConsulta.AppendLine("WHERE        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fEmpresa", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfEmpresa.ReportSource = loObjetoReporte

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
