'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "fCronograma"
'-------------------------------------------------------------------------------------------'
Partial Class fCronograma

    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("--**************************************************************")
            loConsulta.AppendLine("-- Busca los datos del encabezado")
            loConsulta.AppendLine("--**************************************************************")
            loConsulta.AppendLine("CREATE TABLE #tmpCrono( Documento   CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Mes         INT,")
            loConsulta.AppendLine("                        Año         INT,")
            loConsulta.AppendLine("                        Estatus     CHAR(15)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Cod_Cal     CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Nom_Cal     CHAR(100)   COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Cod_Con     CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Nom_Con     CHAR(100)   COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Cod_Dep     CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Nom_Dep     CHAR(100)   COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Cod_Car     CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Nom_Car     CHAR(100)   COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Cod_Tur     CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Nom_Tur     CHAR(100)   COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                        Comentario  VARCHAR(MAX)COLLATE DATABASE_DEFAULT")
            loConsulta.AppendLine("                        );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpCrono(  Documento, MEs, Año, Estatus, Cod_Cal, Nom_Cal, ")
            loConsulta.AppendLine("                        Cod_Con, Nom_Con, Cod_Dep, Nom_Dep,")
            loConsulta.AppendLine("                        Cod_Car, Nom_Car, Cod_Tur, Nom_Tur, Comentario)")
            loConsulta.AppendLine("SELECT  Cronogramas.Documento                       AS Documento, ")
            loConsulta.AppendLine("        Cronogramas.Mes                             AS Mes, ")
            loConsulta.AppendLine("        Cronogramas.Año                             AS Año, ")
            loConsulta.AppendLine("        Cronogramas.Status                          AS Estatus, ")
            loConsulta.AppendLine("        Cronogramas.cod_cal                         AS Cod_cal, ")
            loConsulta.AppendLine("        COALESCE(Calendarios.Nom_Cal, '')           AS Nom_Cal, ")
            loConsulta.AppendLine("        Cronogramas.cod_con                         AS Cod_con, ")
            loConsulta.AppendLine("        COALESCE(Contratos.Nom_Con, '')             AS Nom_Con, ")
            loConsulta.AppendLine("        Cronogramas.cod_dep                         AS Cod_Dep, ")
            loConsulta.AppendLine("        COALESCE(Departamentos_Nomina.Nom_dep, '')  AS Nom_Dep, ")
            loConsulta.AppendLine("        Cronogramas.cod_car                         AS Cod_Car, ")
            loConsulta.AppendLine("        COALESCE(Cargos.Nom_Car, '')                AS Nom_Car, ")
            loConsulta.AppendLine("        Cronogramas.Cod_Tur                         AS Cod_Tur,")
            loConsulta.AppendLine("        COALESCE(Turnos_Nomina.Nom_Tur, '')         AS Nom_Tur,")
            loConsulta.AppendLine("        Cronogramas.Comentario                      AS Comentario")
            loConsulta.AppendLine("FROM    Cronogramas ")
            loConsulta.AppendLine("    LEFT JOIN Calendarios ON Calendarios.Cod_Cal = Cronogramas.Cod_Cal")
            loConsulta.AppendLine("    LEFT JOIN Contratos ON Contratos.Cod_Con = Cronogramas.Cod_Con")
            loConsulta.AppendLine("    LEFT JOIN Departamentos_Nomina ON Departamentos_Nomina.Cod_dep = Cronogramas.Cod_dep")
            loConsulta.AppendLine("    LEFT JOIN Cargos ON Cargos.Cod_Car = Cronogramas.Cod_Car")
            loConsulta.AppendLine("    LEFT JOIN Turnos_Nomina ON Turnos_Nomina.Cod_Tur = Cronogramas.Cod_Tur")
            loConsulta.AppendLine("WHERE        " & cusAplicacion.goFormatos.pcCondicionPrincipal)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--**************************************************************")
            loConsulta.AppendLine("-- Busca los renglones")
            loConsulta.AppendLine("--**************************************************************")
            loConsulta.AppendLine("DECLARE @lcDocumento CHAR(10); ")
            loConsulta.AppendLine("SET @lcDocumento = (SELECT TOP 1 Documento FROM #tmpCrono);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpRenglones( Renglon     INT, ")
            loConsulta.AppendLine("                            Cod_Tra     CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Nom_Tra     CHAR(100)   COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Fecha       DATE, ")
            loConsulta.AppendLine("                            Actividad   CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Revision1   CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Revision2   CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Color_Rev1  CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Color_Rev2  CHAR(10)    COLLATE DATABASE_DEFAULT")
            loConsulta.AppendLine("                            );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpRenglones(  Cod_Tra, Fecha, Actividad, ")
            loConsulta.AppendLine("                            Revision1, Revision2, Color_Rev1, Color_Rev2)")
            loConsulta.AppendLine("SELECT      RC.Cod_Tra,                 RC.Fecha,       RC.Actividad, ")
            loConsulta.AppendLine("            RC.Revision1,               RC.Revision2,   ")
            loConsulta.AppendLine("            COALESCE(S1.val_car,''),    COALESCE(S2.val_car, '')")
            loConsulta.AppendLine("FROM        Renglones_Cronogramas AS RC")
            loConsulta.AppendLine("    JOIN    Trabajadores AS T ")
            loConsulta.AppendLine("        ON  T.Cod_Tra = RC.Cod_Tra")
            loConsulta.AppendLine("    LEFT JOIN renglones_series AS S1")
            loConsulta.AppendLine("        ON  S1.Car_Ini = RC.Revision1")
            loConsulta.AppendLine("        AND S1.Cod_Ser = 'REVCRONOM1'")
            loConsulta.AppendLine("    LEFT JOIN renglones_series AS S2")
            loConsulta.AppendLine("        ON  S2.Car_Ini = RC.Revision2")
            loConsulta.AppendLine("        AND S2.Cod_Ser = 'REVCRONOM2'")
            loConsulta.AppendLine("WHERE       Documento = @lcDocumento;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE CLUSTERED INDEX PK_tmpRenglones_Cod_Tra ON #tmpRenglones(Cod_Tra);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--**************************************************************")
            loConsulta.AppendLine("-- Genera la tabla ""agrupada/resumida"" de actividades")
            loConsulta.AppendLine("--**************************************************************")
            loConsulta.AppendLine("DECLARE @lnAño INT; ")
            loConsulta.AppendLine("DECLARE @lnMes INT; ")
            loConsulta.AppendLine("DECLARE @lcAño CHAR(4); ")
            loConsulta.AppendLine("DECLARE @lcMes CHAR(2); ")
            loConsulta.AppendLine("DECLARE @lnDias INT; ")
            loConsulta.AppendLine("DECLARE @lnContador INT; ")
            loConsulta.AppendLine("DECLARE @ldDesde DATETIME; ")
            loConsulta.AppendLine("DECLARE @ldHasta DATETIME; ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SET @lnAño = (SELECT TOP 1 Año FROM #tmpCrono);")
            loConsulta.AppendLine("SET @lcAño = RIGHT('0000' + CAST(@lnAño AS VARCHAR(4)), 4);")
            loConsulta.AppendLine("SET @lnMes = (SELECT TOP 1 Mes FROM #tmpCrono);")
            loConsulta.AppendLine("SET @lcMes = RIGHT('0000' + CAST(@lnMes AS VARCHAR(2)), 2);")
            loConsulta.AppendLine("SET @ldDesde = (SELECT DATEADD(MONTH, @lnMes-1, DATEADD(YEAR, @lnAño-1900, CAST('19000101' AS DATE))));")
            loConsulta.AppendLine("SET @ldHasta = DATEADD(DAY, -1, DATEADD(MONTH, 1, @ldDesde));")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SET @lnDias = DATEDIFF(DAY, @ldDesde, @ldHasta)+1")
            loConsulta.AppendLine("SET @lnContador = 0;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpDetalles(  Renglon     INT, ")
            loConsulta.AppendLine("                            Cod_Tra     CHAR(10)    COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                            Nom_Tra     CHAR(100)   COLLATE DATABASE_DEFAULT);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpDetalles(Renglon, Cod_Tra, Nom_Tra)")
            loConsulta.AppendLine("SELECT      ROW_NUMBER() OVER (ORDER BY RC.Cod_Tra ASC), RC.Cod_Tra, T.Nom_Tra")
            loConsulta.AppendLine("FROM        Renglones_Cronogramas AS RC")
            loConsulta.AppendLine("    JOIN    Trabajadores AS T")
            loConsulta.AppendLine("        ON  T.Cod_Tra = RC.Cod_Tra")
            loConsulta.AppendLine("WHERE       Documento = @lcDocumento")
            loConsulta.AppendLine("GROUP BY    RC.Cod_Tra, T.Nom_Tra;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE CLUSTERED INDEX PK_tmpDetalles_Cod_Tra ON #tmpDetalles(Cod_Tra);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--Inserta los renglones en el detalle resumido")
            loConsulta.AppendLine("DECLARE @lcConsulta VARCHAR(MAX);")
            loConsulta.AppendLine("DECLARE @lcIndice CHAR(2);")
            loConsulta.AppendLine("WHILE ( @lnContador < 31)")
            loConsulta.AppendLine("BEGIN ")
            loConsulta.AppendLine("    SET @lnContador = @lnContador + 1;")
            loConsulta.AppendLine("    SET @lcIndice = RIGHT('00' + CAST(@lnContador AS VARCHAR(2)), 2);")
            loConsulta.AppendLine("    SET @lcConsulta = 'ALTER TABLE #tmpDetalles ADD [Dia_' + @lcIndice + '] CHAR(1);'  + CHAR(13) +")
            loConsulta.AppendLine("                      'ALTER TABLE #tmpDetalles ADD [Act_' + @lcIndice + '] CHAR(10);' + CHAR(13) +")
            loConsulta.AppendLine("                      'ALTER TABLE #tmpDetalles ADD [Re1_' + @lcIndice + '] CHAR(10);' + CHAR(13) +")
            loConsulta.AppendLine("                      'ALTER TABLE #tmpDetalles ADD [Re2_' + @lcIndice + '] CHAR(10);' + CHAR(13) +")
            loConsulta.AppendLine("                      'ALTER TABLE #tmpDetalles ADD [CRe1_' + @lcIndice + '] CHAR(10);' + CHAR(13) +")
            loConsulta.AppendLine("                      'ALTER TABLE #tmpDetalles ADD [CRe2_' + @lcIndice + '] CHAR(10);' + CHAR(13) +")
            loConsulta.AppendLine("                      '' + CHAR(13) ;")
            loConsulta.AppendLine("    EXECUTE (@lcConsulta);")
            loConsulta.AppendLine("    ")
            loConsulta.AppendLine("    IF (@lnContador <= @lnDias)")
            loConsulta.AppendLine("    BEGIN")
            loConsulta.AppendLine("        SET @lcConsulta = 'UPDATE #tmpDetalles' + CHAR(13) +")
            loConsulta.AppendLine("                          'SET [Dia_' + @lcIndice + '] = R.Dia,' + CHAR(13) +")
            loConsulta.AppendLine("                          '    [Act_' + @lcIndice + '] = R.Actividad,' + CHAR(13) +")
            loConsulta.AppendLine("                          '    [Re1_' + @lcIndice + '] = R.Revision1,' + CHAR(13) +")
            loConsulta.AppendLine("                          '    [Re2_' + @lcIndice + '] = R.Revision2, ' + CHAR(13) +")
            loConsulta.AppendLine("                          '    [CRe1_' + @lcIndice + '] = Color_Rev1, ' + CHAR(13) +")
            loConsulta.AppendLine("                          '    [CRe2_' + @lcIndice + '] = Color_Rev2 ' + CHAR(13) +")
            loConsulta.AppendLine("                          'FROM (   SELECT  T.Cod_Tra, Fecha, Actividad, Revision1, Revision2, Color_Rev1, Color_Rev2,' + CHAR(13) +")
            loConsulta.AppendLine("                          '                 (CASE dbo.[udf_GetISOWeekDay](Fecha)' + CHAR(13) +")
            loConsulta.AppendLine("                          '                     WHEN 1 THEN ''L''' + CHAR(13) +")
            loConsulta.AppendLine("                          '                     WHEN 2 THEN ''M''' + CHAR(13) +")
            loConsulta.AppendLine("                          '                     WHEN 3 THEN ''M''' + CHAR(13) +")
            loConsulta.AppendLine("                          '                     WHEN 4 THEN ''J''' + CHAR(13) +")
            loConsulta.AppendLine("                          '                     WHEN 5 THEN ''V''' + CHAR(13) +")
            loConsulta.AppendLine("                          '                     WHEN 6 THEN ''S''' + CHAR(13) +")
            loConsulta.AppendLine("                          '                     WHEN 7 THEN ''D''' + CHAR(13) +")
            loConsulta.AppendLine("                          '                 END) AS Dia' + CHAR(13) +")
            loConsulta.AppendLine("                          '         FROM    #tmpRenglones AS T' + CHAR(13) +")
            loConsulta.AppendLine("                          '         WHERE   Fecha = ''' + @lcAño + @lcMes + @lcIndice + '''' + CHAR(13) +")
            loConsulta.AppendLine("                          '      ) AS R' + CHAR(13) +")
            loConsulta.AppendLine("                          'WHERE #tmpDetalles.Cod_Tra = R.Cod_Tra' + CHAR(13) +")
            loConsulta.AppendLine("                          '' + CHAR(13) ; ")
            loConsulta.AppendLine("                        ")
            loConsulta.AppendLine("        EXECUTE (@lcConsulta);")
            loConsulta.AppendLine("    END")
            loConsulta.AppendLine("    ")
            loConsulta.AppendLine("    ")
            loConsulta.AppendLine("END")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT  Documento, Mes, Año, Estatus, Cod_Cal, Nom_Cal, ")
            loConsulta.AppendLine("        Cod_Con, Nom_Con, Cod_Dep, Nom_Dep, ")
            loConsulta.AppendLine("        Cod_Car, Nom_Car, Cod_Tur, Nom_Tur, Comentario, ")
            loConsulta.AppendLine("        #tmpDetalles.* ")
            loConsulta.AppendLine("FROM    #tmpCrono")
            loConsulta.AppendLine("    CROSS JOIN #tmpDetalles;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--DROP TABLE #tmpCrono;")
            loConsulta.AppendLine("--DROP TABLE #tmpRenglones;")
            loConsulta.AppendLine("--DROP TABLE #tmpDetalles;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos()

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            Me.mCargarLogoEmpresa(laDatosReporte.Tables(0), "LogoEmpresa")

            loObjetoReporte = cusAplicacion.goFormatos.mCargarInforme("fCronograma", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvfCronograma.ReportSource = loObjetoReporte

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
' RJG: 11/12/13: Programacion inicial
'-------------------------------------------------------------------------------------------'
