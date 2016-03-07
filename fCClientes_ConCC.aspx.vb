'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCClientes_ConCC"
'-------------------------------------------------------------------------------------------'
Partial Class fCClientes_ConCC
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim loConsulta As New StringBuilder()


            loConsulta.AppendLine("")
            loConsulta.AppendLine("--Tabla temporal con los registros a listar")
            loConsulta.AppendLine("CREATE TABLE #tmpRegistros( Codigo VARCHAR(10) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Nombre VARCHAR(100) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Estatus VARCHAR(15) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                            Contable XML")
            loConsulta.AppendLine("                            );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpRegistros(Codigo, Nombre, Estatus, Contable)")
            loConsulta.AppendLine("SELECT  Cod_Cla, ")
            loConsulta.AppendLine("        Nom_Cla,")
            loConsulta.AppendLine("        (CASE Status ")
            loConsulta.AppendLine("            WHEN 'A' THEN 'Activo'  ")
            loConsulta.AppendLine("            WHEN 'I' THEN 'Inactivo'")
            loConsulta.AppendLine("            ELSE 'Suspendido'  ")
            loConsulta.AppendLine("        END) AS Status,")
            loConsulta.AppendLine("        Contable")
            loConsulta.AppendLine("FROM    Clases_Clientes")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- En el SELECT final se expande el XML Contable para obtener las ")
            loConsulta.AppendLine("-- Cuentas Contables, de Gastos y Centros de Costos de cada página del registro ")
            loConsulta.AppendLine("SELECT  #tmpRegistros.Codigo                                AS Codigo,")
            loConsulta.AppendLine("        #tmpRegistros.Nombre                                AS Nombre,")
            loConsulta.AppendLine("        #tmpRegistros.Estatus                               AS Estatus,")
            loConsulta.AppendLine("        COALESCE(Detalles.Numero, 1)                        AS Numero,")
            loConsulta.AppendLine("        COALESCE(Detalles.Pagina, '')                       AS Pagina,")
            loConsulta.AppendLine("        COALESCE(Detalles.Cue_Con_Codigo, '')               AS Cue_Con_Codigo,")
            loConsulta.AppendLine("        COALESCE(Cuentas_Contables.Nom_Cue, '')             AS Cue_Con_Nombre,")
            loConsulta.AppendLine("        COALESCE(Detalles.Cue_Gas_Codigo, '')               AS Cue_Gas_Codigo,")
            loConsulta.AppendLine("        COALESCE(Cuentas_Gastos.Nom_Gas, '')                AS Cue_Gas_Nombre,")
            loConsulta.AppendLine("        COALESCE(Detalles.Cen_Cos_Codigo, '')               AS Cen_Cos_Codigo,")
            loConsulta.AppendLine("        COALESCE(Centros_Costos.Nom_Cen, '')                AS Cen_Cos_Nombre,")
            loConsulta.AppendLine("        COALESCE(Detalles.Cen_Cos_Porcentaje, 0)            AS Cen_Cos_Porcentaje ")
            loConsulta.AppendLine("FROM    #tmpRegistros")
            loConsulta.AppendLine("    LEFT JOIN ( SELECT  Codigo,")
            loConsulta.AppendLine("                        (Ficha.C.value('@n[1]', 'VARCHAR(MAX)')+1) AS Numero,")
            loConsulta.AppendLine("                        Ficha.C.value('@nombre[1]', 'VARCHAR(MAX)') AS Pagina,")
            loConsulta.AppendLine("                        Ficha.C.value('./cue_con[1]', 'VARCHAR(MAX)') AS Cue_Con_Codigo,")
            loConsulta.AppendLine("                        Ficha.C.value('./cue_gas[1]', 'VARCHAR(MAX)') AS Cue_Gas_Codigo,")
            loConsulta.AppendLine("                        Costos.C.value('@codigo[1]', 'VARCHAR(MAX)') AS Cen_Cos_Codigo,")
            loConsulta.AppendLine("                        CAST(Costos.C.value('@porcentaje[1]', 'VARCHAR(MAX)') AS DECIMAL(28,10)) AS Cen_Cos_Porcentaje")
            loConsulta.AppendLine("                FROM    #tmpRegistros")
            loConsulta.AppendLine("                    CROSS APPLY Contable.nodes('contable/ficha') AS Ficha(C)")
            loConsulta.AppendLine("                    OUTER APPLY Contable.nodes('contable/ficha/centro_costo') AS Costos(C)")
            loConsulta.AppendLine("            ) Detalles")
            loConsulta.AppendLine("        ON  Detalles.Codigo = #tmpRegistros.Codigo")
            loConsulta.AppendLine("    LEFT JOIN Cuentas_Contables")
            loConsulta.AppendLine("        ON Cuentas_Contables.Cod_Cue = Detalles.Cue_Con_Codigo")
            loConsulta.AppendLine("    LEFT JOIN Cuentas_Gastos")
            loConsulta.AppendLine("        ON Cuentas_Gastos.Cod_Gas = Detalles.Cue_Gas_Codigo")
            loConsulta.AppendLine("    LEFT JOIN Centros_Costos")
            loConsulta.AppendLine("        ON Centros_Costos.Cod_Cen = Detalles.Cen_Cos_Codigo")
            loConsulta.AppendLine("ORDER BY #tmpRegistros.Codigo, COALESCE(Detalles.Numero, 1)")
            loConsulta.AppendLine("")

            'loConsulta.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCClientes_ConCC", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCClientes_ConCC.ReportSource = loObjetoReporte

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

            Me.WbcAdministradorMensajeModal.mMostrarMensajeModal("Error", _
                          "No se pudo Completar el Proceso: " & loExcepcion.Message, _
                           vis3Controles.wbcAdministradorMensajeModal.enumTipoMensaje.KN_Error, _
                           "auto", _
                           "auto")

        End Try

    End Sub

End Class
'-------------------------------------------------------------------------------------------'
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 31/10/14: Codigo inicial															    '
'-------------------------------------------------------------------------------------------'
' JJD: 31/10/14: Adecuacion para Cajas                                                      '
'-------------------------------------------------------------------------------------------'
