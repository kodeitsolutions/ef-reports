'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "KDE_fOrden_Servicio"
'-------------------------------------------------------------------------------------------'
Partial Class KDE_fOrden_Servicio

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT  Clientes.Cod_Cli                                 AS Cod_Reg, ")
            loConsulta.AppendLine("        Clientes.Nom_Cli                                 AS Nom_Reg, ")
            loConsulta.AppendLine("        Clientes.dir_fis                                 AS Direccion,")
            loConsulta.AppendLine("        Clientes.telefonos                               AS Telefono,")
            loConsulta.AppendLine("        Clientes.correo                                  AS Correo_cli,")
            loConsulta.AppendLine("        COALESCE(Ejec.correo,'')                         AS Correo_ven,")
            loConsulta.AppendLine("        Casos.Documento                                  AS Documento, ")
            loConsulta.AppendLine("        Casos.Status                                     AS Status, ")
            loConsulta.AppendLine("        Casos.Fec_Ini                                    AS Fec_Ini, ")
            loConsulta.AppendLine("        Casos.Fec_Fin                                    AS Fec_Fin, ")
            loConsulta.AppendLine("        (CASE Casos.Status ")
            loConsulta.AppendLine("            WHEN 'Pendiente'  ")
            loConsulta.AppendLine("            THEN DATEDIFF(DAY,Casos.Fec_ini, GETDATE()) ")
            loConsulta.AppendLine("            ELSE DATEDIFF(DAY,Casos.Fec_ini, Casos.Fec_fin) ")
            loConsulta.AppendLine("        END)                                             AS Dias, ")
            loConsulta.AppendLine("        Casos.Asunto                                     AS Asunto, ")
            loConsulta.AppendLine("        Casos.Comentario                                 AS Comentario, ")
            ''loConsulta.AppendLine("        Casos.Cod_Coo                                   AS Cod_Coo, ")
            ''loConsulta.AppendLine("        (CASE WHEN Casos.Cod_Coo<>''")
            ''loConsulta.AppendLine("        THEN(RTRIM(Casos.Cod_Coo) + ' - ' ")
            ''loConsulta.AppendLine("            + COALESCE(Coord.Nom_Ven,'[No Válido]') ) ")
            ''loConsulta.AppendLine("        ELSE '[No Asignado]'")
            ''loConsulta.AppendLine("        END)                                            AS Nom_Coo,")
            loConsulta.AppendLine("        (CASE WHEN Casos.Cod_Eje<>''")
            loConsulta.AppendLine("        THEN(RTRIM(Casos.Cod_Eje) + ' - ' ")
            loConsulta.AppendLine("            + COALESCE(Ejec.Nom_Ven,'[No Válido]') ) ")
            loConsulta.AppendLine("        ELSE '[No Asignado]'")
            loConsulta.AppendLine("        END)                                             AS Nom_Eje, ")
            loConsulta.AppendLine("        Casos.Departamento                               AS Departamento, ")
            loConsulta.AppendLine("        Casos.Origen                                     AS Origen, ")
            loConsulta.AppendLine("        Casos.Principal                                  AS Principal,")
            loConsulta.AppendLine("        COALESCE(Renglones_Casos.Renglon,0)              AS Renglon,")
            loConsulta.AppendLine("        Renglones_Casos.Fec_Ini                          AS Renglon_Fec_Ini,")
            loConsulta.AppendLine("        Renglones_Casos.Hor_Ini                          AS Renglon_Hor_Ini,")
            loConsulta.AppendLine("        Renglones_Casos.Hor_Fin                          AS Renglon_Hor_Fin,")
            loConsulta.AppendLine("        COALESCE(Renglones_Casos.Duracion,0)             AS Renglon_Duracion,")
            loConsulta.AppendLine("        COALESCE(Renglones_Casos.Actividad,'')           AS Renglon_Actividad")
            loConsulta.AppendLine("FROM    Casos")
            loConsulta.AppendLine("    JOIN Clientes ON Clientes.Cod_Cli = Casos.Cod_Reg")
            loConsulta.AppendLine("    LEFT JOIN Vendedores Coord ON Coord.Cod_Ven = Casos.Cod_Coo")
            loConsulta.AppendLine("    LEFT JOIN Vendedores Ejec ON Ejec.Cod_Ven = Casos.Cod_Eje")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Casos ON Renglones_Casos.Documento = Casos.Documento ")
            loConsulta.AppendLine("WHERE    " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("ORDER BY COALESCE(Renglones_Casos.Renglon,0) ASC")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")



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

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("KDE_fOrden_Servicio", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvKDE_fOrden_Servicio.ReportSource = loObjetoReporte

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
' MAVM: 20/04/17: Codigo inicial
'-------------------------------------------------------------------------------------------'
