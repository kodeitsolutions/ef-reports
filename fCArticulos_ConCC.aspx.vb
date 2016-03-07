'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCArticulos_ConCC"
'-------------------------------------------------------------------------------------------'
Partial Class fCArticulos_ConCC
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try


            Dim loComandoSeleccionar As New StringBuilder()


            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Tabla temporal con los registros a listar")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpRegistros( Codigo VARCHAR(10) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Nombre VARCHAR(100) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Estatus VARCHAR(15) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Contable XML")
            loComandoSeleccionar.AppendLine("                            );")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpRegistros(Codigo, Nombre, Estatus, Contable)")
            loComandoSeleccionar.AppendLine("SELECT  Cod_Cla, ")
            loComandoSeleccionar.AppendLine("        Nom_Cla,")
            loComandoSeleccionar.AppendLine("        (CASE Status ")
            loComandoSeleccionar.AppendLine("            WHEN 'A' THEN 'Activo'  ")
            loComandoSeleccionar.AppendLine("            WHEN 'I' THEN 'Inactivo'")
            loComandoSeleccionar.AppendLine("            ELSE 'Suspendido'  ")
            loComandoSeleccionar.AppendLine("        END) AS Status,")
            loComandoSeleccionar.AppendLine("        Contable")
            loComandoSeleccionar.AppendLine("FROM    Clases_Articulos")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- En el SELECT final se expande el XML Contable para obtener las ")
            loComandoSeleccionar.AppendLine("-- Cuentas Contables, de Gastos y Centros de Costos de cada página del registro ")
            loComandoSeleccionar.AppendLine("SELECT  CASE WHEN (LEN(Detalles.Cue_Con_Codigo) > '0' AND (LEN(Detalles.Cue_Con_Codigo) < '9' OR LEN(Detalles.Cue_Con_Codigo) > '9')) THEN '******' ELSE '' END	AS Asteriscos,")
            loComandoSeleccionar.AppendLine("        #tmpRegistros.Codigo                                AS Codigo,")
            loComandoSeleccionar.AppendLine("        #tmpRegistros.Nombre                                AS Nombre,")
            loComandoSeleccionar.AppendLine("        #tmpRegistros.Estatus                               AS Estatus,")
            loComandoSeleccionar.AppendLine("        #tmpRegistros.Estatus                               AS Status_Conceptos, ")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Numero, 1)                        AS Numero,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Pagina, '')                       AS Pagina,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cue_Con_Codigo, '')               AS Cue_Con_Codigo,")
            loComandoSeleccionar.AppendLine("        COALESCE(Cuentas_Contables.Nom_Cue, '')             AS Cue_Con_Nombre,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cue_Gas_Codigo, '')               AS Cue_Gas_Codigo,")
            loComandoSeleccionar.AppendLine("        COALESCE(Cuentas_Gastos.Nom_Gas, '')                AS Cue_Gas_Nombre,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cen_Cos_Codigo, '')               AS Cen_Cos_Codigo,")
            loComandoSeleccionar.AppendLine("        COALESCE(Centros_Costos.Nom_Cen, '')                AS Cen_Cos_Nombre,")
            loComandoSeleccionar.AppendLine("        COALESCE(Detalles.Cen_Cos_Porcentaje, 0)            AS Cen_Cos_Porcentaje ")
            loComandoSeleccionar.AppendLine("FROM    #tmpRegistros")
            loComandoSeleccionar.AppendLine("    LEFT JOIN ( SELECT  Codigo,")
            loComandoSeleccionar.AppendLine("                        (Ficha.C.value('@n[1]', 'VARCHAR(MAX)')+1) AS Numero,")
            loComandoSeleccionar.AppendLine("                        Ficha.C.value('@nombre[1]', 'VARCHAR(MAX)') AS Pagina,")
            loComandoSeleccionar.AppendLine("                        Ficha.C.value('./cue_con[1]', 'VARCHAR(MAX)') AS Cue_Con_Codigo,")
            loComandoSeleccionar.AppendLine("                        Ficha.C.value('./cue_gas[1]', 'VARCHAR(MAX)') AS Cue_Gas_Codigo,")
            loComandoSeleccionar.AppendLine("                        Costos.C.value('@codigo[1]', 'VARCHAR(MAX)') AS Cen_Cos_Codigo,")
            loComandoSeleccionar.AppendLine("                        CAST(Costos.C.value('@porcentaje[1]', 'VARCHAR(MAX)') AS DECIMAL(28,10)) AS Cen_Cos_Porcentaje")
            loComandoSeleccionar.AppendLine("                FROM    #tmpRegistros")
            loComandoSeleccionar.AppendLine("                    CROSS APPLY Contable.nodes('contable/ficha') AS Ficha(C)")
            loComandoSeleccionar.AppendLine("                    OUTER APPLY Contable.nodes('contable/ficha/centro_costo') AS Costos(C)")
            loComandoSeleccionar.AppendLine("            ) Detalles")
            loComandoSeleccionar.AppendLine("        ON  Detalles.Codigo = #tmpRegistros.Codigo")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Cuentas_Contables")
            loComandoSeleccionar.AppendLine("        ON Cuentas_Contables.Cod_Cue = Detalles.Cue_Con_Codigo")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Cuentas_Gastos")
            loComandoSeleccionar.AppendLine("        ON Cuentas_Gastos.Cod_Gas = Detalles.Cue_Gas_Codigo")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Centros_Costos")
            loComandoSeleccionar.AppendLine("        ON Centros_Costos.Cod_Cen = Detalles.Cen_Cos_Codigo")
            'loComandoSeleccionar.AppendLine("WHERE    Detalles.Cue_Con_Codigo <> '' ")
            'loComandoSeleccionar.AppendLine("   AND   LEN(Detalles.Cue_Con_Codigo) <> '9' ")
            loComandoSeleccionar.AppendLine("ORDER BY #tmpRegistros.Codigo, COALESCE(Detalles.Numero, 1)")

            'loConsulta.AppendLine(" WHERE " & cusAplicacion.goFormatos.pcCondicionPrincipal)

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString)

            Dim loServicios As New cusDatos.goDatos
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString, "curReportes")

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


            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCArticulos_ConCC", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)
            Me.mFormatearCamposReporte(loObjetoReporte)
            Me.crvfCArticulos_ConCC.ReportSource = loObjetoReporte

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
