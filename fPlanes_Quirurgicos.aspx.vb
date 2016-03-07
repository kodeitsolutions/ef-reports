'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "fPlanes_Quirurgicos"
'-------------------------------------------------------------------------------------------'
Partial Class fPlanes_Quirurgicos
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT   Agendas.Documento                                                           AS Documento,")
            loConsulta.AppendLine("         YEAR(Agendas.Fec_Ini)                                                       AS Anio,")
            loConsulta.AppendLine("         dbo.udf_GetISOWeek(Agendas.Fec_Ini)                                         AS Semana,")
            loConsulta.AppendLine("         dbo.udf_GetISOWeekDay(Agendas.Fec_Ini)                                      AS DiaSemana,")
            loConsulta.AppendLine("         DATEADD(DAY, 1-(dbo.udf_GetISOWeekDay(Agendas.Fec_Ini)), Agendas.Fec_Ini)   AS Desde,")
            loConsulta.AppendLine("         DATEADD(DAY, 8-(dbo.udf_GetISOWeekDay(Agendas.Fec_Ini)), Agendas.Fec_Ini)   AS Hasta,")
            loConsulta.AppendLine("         Agendas.Fec_Ini                                                             AS Fecha,")
            loConsulta.AppendLine("         Agendas.Hor_Ini                                                             AS Hora,")
            loConsulta.AppendLine("         Agendas.Horas                                                               AS Horas,")
            loConsulta.AppendLine("         COALESCE(Transportes.Nom_Tra, Agendas.Cod_Tra)                              AS Intervencion,")
            loConsulta.AppendLine("         COALESCE(Vendedores.Nom_Ven, Agendas.Cod_Ven)                               AS Medico,")
            loConsulta.AppendLine("         COALESCE(Pacientes.Nom_Cli, Agendas.Cod_Cli)                                AS Paciente,")
            loConsulta.AppendLine("         COALESCE(Seguros.Nom_Cli, Agendas.Cod_Seg)                                  AS Seguro,")
            loConsulta.AppendLine("         CAST(Agendas.Seleccion1 AS BIT)                                             AS HOP,")
            loConsulta.AppendLine("         CAST(Agendas.Seleccion2 AS BIT)                                             AS AMB,")
            loConsulta.AppendLine("         CAST(Agendas.Seleccion3 AS BIT)                                             AS UCI,")
            loConsulta.AppendLine("         Agendas.Status                                                              AS Status,")
            loConsulta.AppendLine("         Agendas.Comentario                                                          AS Comentario")
            loConsulta.AppendLine("FROM     Agendas")
            loConsulta.AppendLine("    LEFT JOIN Transportes ON Transportes.Cod_Tra = Agendas.Cod_Tra")
            loConsulta.AppendLine("    LEFT JOIN Vendedores ON Vendedores.Cod_Ven = Agendas.Cod_Ven")
            loConsulta.AppendLine("    LEFT JOIN Clientes Pacientes ON Pacientes.Cod_Cli = Agendas.Cod_Cli")
            loConsulta.AppendLine("    LEFT JOIN Clientes Seguros ON Seguros.Cod_Cli = Agendas.Cod_Seg")
            loConsulta.AppendLine("WHERE    Agendas.Status = 'Pendiente' ")
            loConsulta.AppendLine("ORDER BY Agendas.Fec_Ini ASC, Agendas.Hor_Ini ASC")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
           
            Dim loServicios As New cusDatos.goDatos()
            
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")
            
            '--------------------------------------------------'
			' Carga la imagen del logo en cusReportes            '
			'--------------------------------------------------'
			Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")
			
			 '-------------------------------------------------------------------------------------------------------
            ' Verificando si el select (tabla nº0) trae registros
            '-------------------------------------------------------------------------------------------------------

            If (laDatosReporte.Tables(0).Rows.Count <= 0) Then
                Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Información", _
                                          "No se Encontraron Registros para los Parámetros Especificados. ", _
                                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Informacion, _
                                           "350px", _
                                           "200px")
            End If


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fPlanes_Quirurgicos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfPlanes_Quirurgicos.ReportSource = loObjetoReporte

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
' Fin del codigo                                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 17/03/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
