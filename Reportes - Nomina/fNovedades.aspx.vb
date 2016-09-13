'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fNovedades"
'-------------------------------------------------------------------------------------------'
Partial Class fNovedades

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Novedades.Documento             AS documento, ")
            loConsulta.AppendLine("            Novedades.fecha                 AS fecha, ")
            loConsulta.AppendLine("            Novedades.Status                AS status, ")
            loConsulta.AppendLine("            Novedades.cod_tra               AS cod_tra, ")
            loConsulta.AppendLine("            Trabajadores.nom_tra            AS nom_tra,")
            loConsulta.AppendLine("            Novedades.fec_ini               AS fec_ini,")
            loConsulta.AppendLine("            Novedades.fec_fin               AS fec_fin,")
            loConsulta.AppendLine("            Novedades.horas                 AS horas,")
            loConsulta.AppendLine("            Novedades.remunerado            AS remunerado,")
            loConsulta.AppendLine("            Novedades.justificado           AS justificado,")
            loConsulta.AppendLine("            Novedades.aut_por               AS cod_aut,")
            loConsulta.AppendLine("            Autorizacion.nom_tra            AS nom_aut,")
            loConsulta.AppendLine("            Novedades.motivo                AS cod_mot,")
            loConsulta.AppendLine("            Motivos.Nom_mot                 AS nom_mot,")
            loConsulta.AppendLine("            Novedades.concepto              AS cod_con,")
            loConsulta.AppendLine("            Conceptos_Nomina.nom_con        AS nom_con,")
            loConsulta.AppendLine("            Novedades.Tipo                  AS tipo,")
            loConsulta.AppendLine("            Novedades.Clase                 AS clase,")
            loConsulta.AppendLine("            Novedades.Comentario            AS comentario")
            loConsulta.AppendLine("FROM        Novedades")
            loConsulta.AppendLine("    JOIN    Trabajadores ")
            loConsulta.AppendLine("        ON  Trabajadores.cod_tra = Novedades.cod_tra")
            loConsulta.AppendLine("    JOIN    Trabajadores AS Autorizacion ")
            loConsulta.AppendLine("        ON  Autorizacion.cod_tra = Novedades.cod_tra")
            loConsulta.AppendLine("    JOIN    Conceptos_Nomina ")
            loConsulta.AppendLine("        ON  Conceptos_Nomina.cod_con = Novedades.concepto")
            loConsulta.AppendLine("    JOIN    Motivos ")
            loConsulta.AppendLine("        ON  Motivos.cod_mot = Novedades.motivo")
            loConsulta.AppendLine("WHERE        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fNovedades", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfNovedades.ReportSource = loObjetoReporte

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
