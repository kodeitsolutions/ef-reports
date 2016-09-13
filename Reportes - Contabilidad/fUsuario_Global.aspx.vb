'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fUsuario_Global"
'-------------------------------------------------------------------------------------------'
Partial Class fUsuario_Global

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT	Factory_Global.dbo.Usuarios.Cod_usu,")
            loConsulta.AppendLine("         (CASE   Factory_Global.dbo.Usuarios.status")
            loConsulta.AppendLine("		            WHEN 'A' THEN 'ACTIVO'")
            loConsulta.AppendLine("		            WHEN 'I' THEN 'INACTIVO'")
            loConsulta.AppendLine("		            WHEN 'S' THEN 'SUSPENDIDO'")
            loConsulta.AppendLine("		    END) AS status,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.Nom_Usu,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.Rif,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.Tipo,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.clase,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.Departamento,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.Cargo,")
            loConsulta.AppendLine("		    Factory_Global.dbo.idiomas.nom_idi,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.Correo,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.Correo2,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.Telefonos,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.movil,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.direccion,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.nivel,")
            loConsulta.AppendLine("         (CASE   Factory_Global.dbo.Usuarios.Sexo")
            loConsulta.AppendLine("		            WHEN 'F' THEN 'FEMENINO'")
            loConsulta.AppendLine("		            WHEN 'M' THEN 'MASCULINO'")
            loConsulta.AppendLine("		            ELSE ''")
            loConsulta.AppendLine("		    END) AS Sexo,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.fec_nac,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.fec_ini,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.Opc_def,")
            loConsulta.AppendLine("		    Factory_Global.dbo.Usuarios.Comentario")
            loConsulta.AppendLine("FROM     Factory_Global.dbo.Usuarios")
            loConsulta.AppendLine("JOIN     Factory_Global.dbo.idiomas ON Factory_Global.dbo.Idiomas.cod_idi = Factory_Global.dbo.Usuarios.cod_idi")
            loConsulta.AppendLine("WHERE        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fUsuario_Global", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfUsuario_Global.ReportSource = loObjetoReporte

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
