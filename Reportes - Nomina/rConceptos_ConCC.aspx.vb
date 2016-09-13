'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data
'-------------------------------------------------------------------------------------------'
' Inicio de clase "rConceptos_ConCC"
'-------------------------------------------------------------------------------------------'
Partial Class rConceptos_ConCC
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))

        Try

            Dim loComandoSeleccionar As New StringBuilder()

            'loConsulta.AppendLine("SELECT  Cod_Con,")
            'loConsulta.AppendLine("        Nom_Con,")
            'loConsulta.AppendLine("        Status,")
            'loConsulta.AppendLine("        (CASE WHEN Status = 'A' THEN 'Activo' ELSE 'Inactivo' END) AS Status_Conceptos ")
            'loConsulta.AppendLine("FROM    Conceptos")
            'loConsulta.AppendLine("WHERE   Cod_Con BETWEEN " & lcParametro0Desde)
            'loConsulta.AppendLine("    AND " & lcParametro0Hasta)
            'loConsulta.AppendLine("    AND Status IN (" & lcParametro1Desde & ")")
            'loConsulta.AppendLine("ORDER BY Cod_Con, Nom_Con")




            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("--Tabla temporal con los registros a listar")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpRegistros( Codigo VARCHAR(30) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Nombre VARCHAR(100) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Estatus VARCHAR(15) COLLATE DATABASE_DEFAULT, ")
            loComandoSeleccionar.AppendLine("                            Contable XML")
            loComandoSeleccionar.AppendLine("                            );")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpRegistros(Codigo, Nombre, Estatus, Contable)")
            loComandoSeleccionar.AppendLine("SELECT  Cod_Con, ")
            loComandoSeleccionar.AppendLine("        Nom_Con,")
            loComandoSeleccionar.AppendLine("        (CASE Status ")
            loComandoSeleccionar.AppendLine("            WHEN 'A' THEN 'Activo'  ")
            loComandoSeleccionar.AppendLine("            WHEN 'I' THEN 'Inactivo'")
            loComandoSeleccionar.AppendLine("            ELSE 'Suspendido'  ")
            loComandoSeleccionar.AppendLine("        END) AS Status,")
            loComandoSeleccionar.AppendLine("        Contable")
            loComandoSeleccionar.AppendLine("FROM    Conceptos_Nomina")
            loComandoSeleccionar.AppendLine("WHERE   Conceptos_Nomina.Cod_Con   Between " & lcParametro0Desde)
            loComandoSeleccionar.AppendLine("        And " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("        And Conceptos_Nomina.Status        IN (" & lcParametro1Desde & ")")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("-- En el SELECT final se expande el XML Contable para obtener las ")
            loComandoSeleccionar.AppendLine("-- Cuentas Contables, de Gastos y Centros de Costos de cada página del registro ")
            loComandoSeleccionar.AppendLine("")
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
            loComandoSeleccionar.AppendLine("")


            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rConceptos_ConCC", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrConceptos_ConCC.ReportSource = loObjetoReporte

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
' MJP: 09/07/08: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' MJP: 11/07/08: Creación objeto que cierra el archivo de reporte.                          '
'-------------------------------------------------------------------------------------------'
' MJP: 14/07/08: Agregacion filtro Status.                                                  '
'-------------------------------------------------------------------------------------------'
' MVP: 04/08/08: Cambios para multi idioma, mensaje de error y clase padre.                 '
'-------------------------------------------------------------------------------------------'
' RJG: 10/09/14: Ajuste de interfaz y estandarización de código.                            '
'-------------------------------------------------------------------------------------------'
' JJD: 04/12/14: Programacion de la busqueda de las Cuentas Contables.                      '
'-------------------------------------------------------------------------------------------'
' JJD: 09/12/14: Ajuste a la emision del reporte.                                           '
'-------------------------------------------------------------------------------------------'
' JJD: 18/12/14: Inclusion del Len de la Cuenta Contable.                                   '
'-------------------------------------------------------------------------------------------'
' EAG: 30/09/15: Comentar linea que escribe la consulta SQL.                                '
'-------------------------------------------------------------------------------------------'
