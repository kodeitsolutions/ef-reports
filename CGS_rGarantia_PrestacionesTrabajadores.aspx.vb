﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "CGS_rGarantia_PrestacionesTrabajadores"
'-------------------------------------------------------------------------------------------'
Partial Class CGS_rGarantia_PrestacionesTrabajadores
    Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument
	
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
        Dim lcParametro3Desde As String = cusAplicacion.goReportes.paParametrosIniciales(3)

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            loConsulta.AppendLine("DECLARE @lcCodTra_Desde AS VARCHAR(12) = " & lcParametro0Desde)
            loConsulta.AppendLine("DECLARE @lcCodTra_Hasta AS VARCHAR(12) = " & lcParametro0Hasta)
            loConsulta.AppendLine("DECLARE @ldFechaFin DATE;")
            loConsulta.AppendLine("SET @ldFechaFin = CAST(" & lcParametro2Desde & " AS DATE);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Obtiene los trabajadores ")
            loConsulta.AppendLine("CREATE TABLE #tmpTrabajadores(  Cod_Tra CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Nom_Tra CHAR(100) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Status CHAR(1) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Cedula CHAR(30) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Ingreso DATE,")
            loConsulta.AppendLine("                                Egreso DATE,")
            loConsulta.AppendLine("                                Antiguedad INT,")
            loConsulta.AppendLine("                                Meses_Antiguedad INT,")
            loConsulta.AppendLine("                                Anos_Antiguedad INT,")
            loConsulta.AppendLine("                                Anos_Prestaciones INT,")
            loConsulta.AppendLine("                                Dias_Prestaciones_Mes INT,")
            loConsulta.AppendLine("                                Dias_Prestaciones_Ano INT,")
            loConsulta.AppendLine("                                Dias_Ultimo_Sueldo INT,")
            loConsulta.AppendLine("                                Prestaciones_Ultimo_Sueldo DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Ultimo_Sueldo_Mensual DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Ultimo_Sueldo_Diario DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Inicial_Prestaciones DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Inicial_Anticipo_Prestaciones DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Inicial_Intereses_Prestaciones DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Inicial_Dias_Prestaciones DECIMAL(28, 10)")
            loConsulta.AppendLine("                            );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpTrabajadores(Cod_Tra, Nom_Tra, Status, Cedula, Ingreso, Egreso, Ultimo_Sueldo_Mensual,")
            loConsulta.AppendLine("                             Inicial_Prestaciones, Inicial_Anticipo_Prestaciones,")
            loConsulta.AppendLine("                             Inicial_Intereses_Prestaciones, Inicial_Dias_Prestaciones)")
            loConsulta.AppendLine("SELECT      Trabajadores.Cod_Tra                                    AS Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                                    AS Nom_Tra,")
            loConsulta.AppendLine("            Trabajadores.Status                                     AS Status,")
            loConsulta.AppendLine("            Trabajadores.Cedula                                     AS Cedula,")
            loConsulta.AppendLine("            Trabajadores.Fec_Ini                                    AS Ingreso,")
            loConsulta.AppendLine("            CASE WHEN Trabajadores.Status = 'L' ")
            loConsulta.AppendLine("                THEN COALESCE(Liquidaciones.Fecha, Trabajadores.Fec_Fin)")
            loConsulta.AppendLine("                ELSE @ldFechaFin")
            loConsulta.AppendLine("            END                                                     AS Egreso,")
            loConsulta.AppendLine("            COALESCE(Sueldo_Mensual.Val_Num, 0)                     AS Ultimo_Sueldo_Mensual,")
            loConsulta.AppendLine("            COALESCE(I_Prestaciones.Val_Num, 0)                     AS Inicial_Prestaciones,")
            loConsulta.AppendLine("            COALESCE(I_Anticipo_Prestaciones.Val_Num, 0)            AS Inicial_Anticipo_Prestaciones,")
            loConsulta.AppendLine("            COALESCE(I_Intereses_Prestaciones.Val_Num, 0)           AS Inicial_Intereses_Prestaciones,")
            loConsulta.AppendLine("            COALESCE(I_Dias_Prestaciones.Val_Num, 0)                AS Inicial_Dias_Prestaciones")
            loConsulta.AppendLine("FROM        Trabajadores")
            loConsulta.AppendLine("    LEFT JOIN Liquidaciones")
            loConsulta.AppendLine("        ON  Liquidaciones.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina I_Prestaciones")
            loConsulta.AppendLine("        ON  I_Prestaciones.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND I_Prestaciones.Cod_Cam = 'Z002'")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina I_Anticipo_Prestaciones")
            loConsulta.AppendLine("        ON  I_Anticipo_Prestaciones.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND I_Anticipo_Prestaciones.Cod_Cam = 'Z004'")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina I_Intereses_Prestaciones")
            loConsulta.AppendLine("        ON  I_Intereses_Prestaciones.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND I_Intereses_Prestaciones.Cod_Cam = 'Z003'")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina I_Dias_Prestaciones")
            loConsulta.AppendLine("        ON  I_Dias_Prestaciones.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND I_Dias_Prestaciones.Cod_Cam = 'Z012'")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina Sueldo_Mensual")
            loConsulta.AppendLine("        ON  Sueldo_Mensual.Cod_Tra = Trabajadores.Cod_Tra")
            loConsulta.AppendLine("        AND Sueldo_Mensual.Cod_Cam = 'A001'")
            loConsulta.AppendLine("WHERE   Trabajadores.Cod_Tra BETWEEN @lcCodTra_Desde AND @lcCodTra_Hasta")
            loConsulta.AppendLine("    AND Trabajadores.Status IN ( " & lcParametro1Desde & " )")
            loConsulta.AppendLine("    AND Trabajadores.Tip_Tra = 'Trabajador'")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Luego los Conceptos de prestaciones")
            loConsulta.AppendLine("CREATE TABLE #tmpPrestaciones(  Orden INTEGER,")
            loConsulta.AppendLine("                                Mes INTEGER,")
            loConsulta.AppendLine("                                Anio INTEGER,")
            loConsulta.AppendLine("                                Cod_Tra CHAR(10) COLLATE DATABASE_DEFAULT,")
            loConsulta.AppendLine("                                Sueldo_Mensual DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Sueldo_Diario DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Ali_Utilidades DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Ali_Bono_Vacacional DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Sueldo_Diario_Integral DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Dias_Prestaciones DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Abono_Prestaciones DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Anticipo_Prestaciones DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Porcentaje_Interes DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Interes_Prestaciones DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Interes_Pagado DECIMAL(28, 10),")
            loConsulta.AppendLine("                            );")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Busca los movimientos de prestaciones de los trabajadores (fondo de garantía)")
            loConsulta.AppendLine("INSERT INTO #tmpPrestaciones(Orden, Mes, Anio, Cod_Tra,")
            loConsulta.AppendLine("        Sueldo_Mensual, Sueldo_Diario, Ali_Utilidades, Ali_Bono_Vacacional,")
            loConsulta.AppendLine("        Sueldo_Diario_Integral, Dias_Prestaciones, Abono_Prestaciones,")
            loConsulta.AppendLine("        Anticipo_Prestaciones, Porcentaje_Interes, Interes_Prestaciones, Interes_Pagado)")
            loConsulta.AppendLine("SELECT  ROW_NUMBER() OVER (  ")
            loConsulta.AppendLine("                PARTITION BY Recibos.Cod_Tra  ")
            loConsulta.AppendLine("                ORDER BY YEAR(Recibos.Fecha), MONTH(Recibos.Fecha)")
            loConsulta.AppendLine("            )                                                       AS Orden,")
            loConsulta.AppendLine("        MONTH(Recibos.Fecha)                                        AS Mes,")
            loConsulta.AppendLine("        YEAR(Recibos.Fecha)                                         AS Anio,")
            loConsulta.AppendLine("        Recibos.Cod_Tra                                             AS Cod_Tra,")
            loConsulta.AppendLine("        SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'O461')")
            loConsulta.AppendLine("            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("            ELSE 0")
            loConsulta.AppendLine("        END)                                                        AS Sueldo_Mensual,")
            loConsulta.AppendLine("        SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'O451')")
            loConsulta.AppendLine("            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("            ELSE 0")
            loConsulta.AppendLine("        END)                                                        AS Sueldo_Diario,")
            loConsulta.AppendLine("        SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'O452')")
            loConsulta.AppendLine("            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("            ELSE 0")
            loConsulta.AppendLine("        END)                                                        AS Ali_Utilidades,")
            loConsulta.AppendLine("        SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'O453')")
            loConsulta.AppendLine("            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("            ELSE 0")
            loConsulta.AppendLine("        END)                                                        AS Ali_Bono_Vacacional,")
            loConsulta.AppendLine("        SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'O450')")
            loConsulta.AppendLine("            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("            ELSE 0")
            loConsulta.AppendLine("        END)                                                        AS Sueldo_Diario_Integral,")
            loConsulta.AppendLine("        SUM(CASE WHEN (Renglones_Recibos.Cod_Con IN ('Y440', 'Y441', 'Y450', 'Y451'))")
            loConsulta.AppendLine("            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("            ELSE 0")
            loConsulta.AppendLine("        END)                                                        AS Dias_Prestaciones,")
            loConsulta.AppendLine("        SUM(CASE WHEN (Renglones_Recibos.Cod_Con IN ('A450', 'A451'))")
            loConsulta.AppendLine("            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("            ELSE 0")
            loConsulta.AppendLine("        END)                                                        AS Abono_Prestaciones,")
            loConsulta.AppendLine("        SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'B009')")
            loConsulta.AppendLine("            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("            ELSE 0")
            loConsulta.AppendLine("        END)                                                        AS Anticipo_Prestaciones,")
            loConsulta.AppendLine("        SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'Y110')")
            loConsulta.AppendLine("            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("            ELSE 0")
            loConsulta.AppendLine("        END)                                                        AS Porcentaje_Interes,")
            loConsulta.AppendLine("        SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'A010')")
            loConsulta.AppendLine("            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("            ELSE 0")
            loConsulta.AppendLine("        END)                                                        AS Interes_Prestaciones,")
            loConsulta.AppendLine("        SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'A101')")
            loConsulta.AppendLine("            THEN Renglones_Recibos.Mon_Net")
            loConsulta.AppendLine("            ELSE 0")
            loConsulta.AppendLine("        END)                                                        AS Interes_Pagado")
            loConsulta.AppendLine("FROM    Recibos")
            loConsulta.AppendLine("    JOIN #tmpTrabajadores ON #tmpTrabajadores.Cod_Tra = Recibos.Cod_Tra")
            loConsulta.AppendLine("    JOIN Renglones_Recibos ON Renglones_Recibos.Documento = Recibos.Documento")
            loConsulta.AppendLine("WHERE   Recibos.Status = 'Confirmado'")
            loConsulta.AppendLine("    AND Recibos.Fecha >= #tmpTrabajadores.Ingreso" )
            loConsulta.AppendLine("    AND Recibos.Fecha <= " & lcParametro2Desde)
            loConsulta.AppendLine("GROUP BY Recibos.Cod_Tra, YEAR(Recibos.Fecha), MONTH(Recibos.Fecha)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Calcula Los meses y años de antiguedad usados para prestaciones a último sueldo")
            loConsulta.AppendLine("-- y los días de prestaciones a último Sueldo ")
            loConsulta.AppendLine("UPDATE  #tmpTrabajadores")
            loConsulta.AppendLine("SET     Dias_Prestaciones_Mes = COALESCE((SELECT TOP 1 Val_Num FROM Constantes_Locales WHERE cod_con = 'C008'), 5),")
            loConsulta.AppendLine("        Dias_Prestaciones_Ano = COALESCE((SELECT TOP 1 Val_Num FROM Constantes_Locales WHERE cod_con = 'C005'), 30);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("UPDATE  #tmpTrabajadores ")
            loConsulta.AppendLine("SET     Antiguedad = DATEDIFF(MONTH, DATEADD(DAY, -DAY(Ingreso)+1, Ingreso ), DATEADD(DAY, -DAY(Ingreso)+1+1, Egreso ) );")
            loConsulta.AppendLine("        ")
            loConsulta.AppendLine("UPDATE  #tmpTrabajadores ")
            loConsulta.AppendLine("SET     Meses_Antiguedad = Antiguedad % 12,")
            loConsulta.AppendLine("        Anos_Antiguedad =  FLOOR( Antiguedad /12.0),")
            loConsulta.AppendLine("        Anos_Prestaciones = FLOOR( Antiguedad /12.0) + (CASE WHEN (Antiguedad % 12>=6) THEN 1 ELSE 0 END);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Obtiene el último sueldo para prestaciones")
            loConsulta.AppendLine("UPDATE      #tmpTrabajadores")
            loConsulta.AppendLine("SET         Ultimo_Sueldo_Diario = Sueldos.Ultimo_Sueldo_Diario")
            loConsulta.AppendLine("FROM    (")
            loConsulta.AppendLine("            SELECT      #tmpTrabajadores.Cod_Tra,")
            loConsulta.AppendLine("                        COALESCE(Ultimos_Sueldos.Sueldo_Diario_Integral, ")
            loConsulta.AppendLine("                            #tmpTrabajadores.Ultimo_Sueldo_Mensual/30.0) AS Ultimo_Sueldo_Diario")
            loConsulta.AppendLine("            FROM        #tmpTrabajadores")
            loConsulta.AppendLine("                LEFT JOIN (")
            loConsulta.AppendLine("                            SELECT  #tmpPrestaciones.Cod_Tra,")
            loConsulta.AppendLine("                                    #tmpPrestaciones.Sueldo_Diario_Integral")
            loConsulta.AppendLine("                            FROM    #tmpPrestaciones")
            loConsulta.AppendLine("                            JOIN ")
            loConsulta.AppendLine("                                (   SELECT  Cod_Tra, MAX(Orden) Orden")
            loConsulta.AppendLine("                                    FROM    #tmpPrestaciones")
            loConsulta.AppendLine("                                    WHERE   Sueldo_Diario_Integral > 0")
            loConsulta.AppendLine("                                    GROUP BY Cod_Tra")
            loConsulta.AppendLine("                                ) X")
            loConsulta.AppendLine("                                ON X.Cod_Tra = #tmpPrestaciones.Cod_Tra")
            loConsulta.AppendLine("                                AND X.Orden = #tmpPrestaciones.Orden")
            loConsulta.AppendLine("                        ) Ultimos_Sueldos")
            loConsulta.AppendLine("                    ON Ultimos_Sueldos.Cod_Tra = #tmpTrabajadores.Cod_Tra")
            loConsulta.AppendLine("        ) Sueldos")
            loConsulta.AppendLine("WHERE   #tmpTrabajadores.Cod_Tra = Sueldos.Cod_Tra")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("--Calcula los días totales y prestaciones al último sueldo")
            loConsulta.AppendLine("UPDATE  #tmpTrabajadores")
            loConsulta.AppendLine("SET     Dias_Ultimo_Sueldo = (CASE  WHEN (Antiguedad < 3) ")
            loConsulta.AppendLine("                                    THEN (Antiguedad+1)*Dias_Prestaciones_Mes ")
            loConsulta.AppendLine("                                    ELSE Anos_Prestaciones*Dias_Prestaciones_Ano")
            loConsulta.AppendLine("                              END),")
            loConsulta.AppendLine("        Prestaciones_Ultimo_Sueldo = (CASE  WHEN (Antiguedad < 3) ")
            loConsulta.AppendLine("                                            THEN (Antiguedad+1)*Dias_Prestaciones_Mes ")
            loConsulta.AppendLine("                                            ELSE (Anos_Prestaciones)*Dias_Prestaciones_Ano")
            loConsulta.AppendLine("                                      END) * Ultimo_Sueldo_Diario;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- SELECT Final")
            loConsulta.AppendLine("SELECT  #tmpTrabajadores.Cod_Tra,")
            loConsulta.AppendLine("        #tmpTrabajadores.Nom_Tra,")
            loConsulta.AppendLine("        #tmpTrabajadores.Status,")
            loConsulta.AppendLine("        (CASE #tmpTrabajadores.Status")
            loConsulta.AppendLine("            WHEN 'A' THEN 'Activo'")
            loConsulta.AppendLine("            WHEN 'S' THEN 'Suspendido'")
            loConsulta.AppendLine("            WHEN 'L' THEN 'Liquidado'")
            loConsulta.AppendLine("            ELSE 'Inactivo'")
            loConsulta.AppendLine("        END) AS Estatus_Trabajador,")
            loConsulta.AppendLine("        #tmpTrabajadores.Cedula,")
            loConsulta.AppendLine("        #tmpTrabajadores.Ingreso,")
            loConsulta.AppendLine("        #tmpTrabajadores.Egreso,")
            loConsulta.AppendLine("        #tmpTrabajadores.Inicial_Prestaciones,")
            loConsulta.AppendLine("        #tmpTrabajadores.Inicial_Anticipo_Prestaciones,")
            loConsulta.AppendLine("        #tmpTrabajadores.Inicial_Intereses_Prestaciones,")
            loConsulta.AppendLine("        #tmpTrabajadores.Inicial_Dias_Prestaciones,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Orden, 1)                  AS Orden,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Mes, MONTH(@ldFechaFin))                    AS Mes,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Anio, YEAR(@ldFechaFin))                   AS Anio,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Sueldo_Mensual, 0)         AS Sueldo_Mensual,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Sueldo_Diario, 0)          AS Sueldo_Diario,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Ali_Utilidades, 0)         AS Ali_Utilidades,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Ali_Bono_Vacacional, 0)    AS Ali_Bono_Vacacional,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Sueldo_Diario_Integral, 0) AS Sueldo_Diario_Integral,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Dias_Prestaciones, 0)      AS Dias_Prestaciones,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Abono_Prestaciones, 0)     AS Abono_Prestaciones,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Anticipo_Prestaciones, 0)  AS Anticipo_Prestaciones,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Porcentaje_Interes, 0)     AS Porcentaje_Interes,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Interes_Prestaciones, 0)   AS Interes_Prestaciones,")
            loConsulta.AppendLine("        COALESCE(#tmpPrestaciones.Interes_Pagado, 0)         AS Interes_Pagado,")
            loConsulta.AppendLine("        #tmpTrabajadores.Meses_Antiguedad,")
            loConsulta.AppendLine("        #tmpTrabajadores.Anos_Antiguedad,")
            loConsulta.AppendLine("        #tmpTrabajadores.Ultimo_Sueldo_Mensual,")
            loConsulta.AppendLine("        #tmpTrabajadores.Dias_Ultimo_Sueldo,")
            loConsulta.AppendLine("        #tmpTrabajadores.Prestaciones_Ultimo_Sueldo")
            loConsulta.AppendLine("FROM    #tmpTrabajadores")
            IF (lcParametro3Desde.Trim().ToUpper() = "SI") Then
                loConsulta.AppendLine("    LEFT JOIN #tmpPrestaciones ON #tmpPrestaciones.Cod_Tra = #tmpTrabajadores.Cod_Tra")
            ELSE
                loConsulta.AppendLine("    JOIN #tmpPrestaciones ON #tmpPrestaciones.Cod_Tra = #tmpTrabajadores.Cod_Tra")
            End If
            loConsulta.AppendLine("ORDER BY #tmpTrabajadores.Cod_Tra,#tmpPrestaciones.Orden ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")


            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("CGS_rGarantia_PrestacionesTrabajadores", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvCGS_rGarantia_PrestacionesTrabajadores.ReportSource = loObjetoReporte

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
' Fin del codigo.                                                                           '
'-------------------------------------------------------------------------------------------'
' RJG: 25/09/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
' RJG: 16/05/15: Se agregaron los conceptos de prestaciones (A451) y días (Y441 y Y451) de  '
'                prestaciones mensuales (viejo esquema).                                    '
'-------------------------------------------------------------------------------------------'
' RJG: 26/05/15: Se amplió el campo Cédula de la tabla temporal a CHAR(30).                 '
'-------------------------------------------------------------------------------------------'