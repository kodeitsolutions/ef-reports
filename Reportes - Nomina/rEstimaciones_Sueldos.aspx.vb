'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rEstimaciones_Sueldos"
'-------------------------------------------------------------------------------------------'
Partial Class rEstimaciones_Sueldos
    Inherits vis2Formularios.frmReporte
	
	Dim loObjetoReporte as CrystalDecisions.CrystalReports.Engine.ReportDocument
	
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = cusAplicacion.goReportes.paParametrosIniciales(2)
        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try

            loConsulta.AppendLine("DECLARE @ldFechaFin DATE; ")
            loConsulta.AppendLine("SET @ldFechaFin = GETDATE(); ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("DECLARE @llAumentarPorcentaje BIT; ")
            loConsulta.AppendLine("DECLARE @lnAumentoPorcentaje DECIMAL(28,10); ")
            loConsulta.AppendLine("DECLARE @lnAumentoMonto DECIMAL(28,10); ")
            loConsulta.AppendLine("")
            IF lcParametro2Desde.Trim().ToUpper() = "PORCENTAJE" Then
                loConsulta.AppendLine("SET @llAumentarPorcentaje = 1; ")
                loConsulta.AppendLine("SET @lnAumentoPorcentaje = " & lcParametro3Desde & "; ")
                loConsulta.AppendLine("SET @lnAumentoMonto = 0; ")
            Else
                loConsulta.AppendLine("SET @llAumentarPorcentaje = 0; ")
                loConsulta.AppendLine("SET @lnAumentoPorcentaje = 0; ")
                loConsulta.AppendLine("SET @lnAumentoMonto = " & lcParametro3Desde & "; ")
            End If
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Obtiene los trabajadores  ")
            loConsulta.AppendLine("CREATE TABLE #tmpTrabajadores(  Cod_Tra CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                                Nom_Tra CHAR(100) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                                Status CHAR(1) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                                Cedula CHAR(10) COLLATE DATABASE_DEFAULT, ")
            loConsulta.AppendLine("                                Ingreso DATE, ")
            loConsulta.AppendLine("                                Egreso DATE, ")
            loConsulta.AppendLine("                                Antiguedad INT, ")
            loConsulta.AppendLine("                                Meses_Antiguedad INT, ")
            loConsulta.AppendLine("                                Anos_Antiguedad INT, ")
            loConsulta.AppendLine("                                Anos_Prestaciones INT, ")
            loConsulta.AppendLine("                                Dias_Prestaciones_Mes INT, ")
            loConsulta.AppendLine("                                Dias_Prestaciones_Ano INT, ")
            loConsulta.AppendLine("                                Prestaciones_Ultimo_Sueldo DECIMAL(28, 10), ")
            loConsulta.AppendLine("                                Ultimo_Sueldo_Mensual DECIMAL(28, 10), ")
            loConsulta.AppendLine("                                Ultimo_Sueldo_Diario DECIMAL(28, 10), ")
            loConsulta.AppendLine("                                Inicial_Prestaciones DECIMAL(28, 10), ")
            loConsulta.AppendLine("                                Inicial_Anticipo_Prestaciones DECIMAL(28, 10), ")
            loConsulta.AppendLine("                                Inicial_Intereses_Prestaciones DECIMAL(28, 10), ")
            loConsulta.AppendLine("                                Inicial_Dias_Prestaciones DECIMAL(28, 10), ")
            loConsulta.AppendLine("                                Porcentaje_Aumento DECIMAL(28, 10), ")
            loConsulta.AppendLine("                                Monto_Aumento DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Nuevo_Sueldo DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Prestaciones_Nuevo_Sueldo DECIMAL(28, 10),")
            loConsulta.AppendLine("                                Diferencia_Pre_Nuevo_Sueldo DECIMAL(28, 10)")
            loConsulta.AppendLine("                            ); ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpTrabajadores(Cod_Tra, Nom_Tra, Status, Cedula, Ingreso, Egreso, ")
            loConsulta.AppendLine("                             Ultimo_Sueldo_Mensual, Ultimo_Sueldo_Diario,")
            loConsulta.AppendLine("                             Inicial_Prestaciones, Inicial_Anticipo_Prestaciones, ")
            loConsulta.AppendLine("                             Inicial_Intereses_Prestaciones, Inicial_Dias_Prestaciones,")
            loConsulta.AppendLine("                             Porcentaje_Aumento, Monto_Aumento, Nuevo_Sueldo) ")
            loConsulta.AppendLine("SELECT      Trabajadores.Cod_Tra                                    AS Cod_Tra, ")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                                    AS Nom_Tra, ")
            loConsulta.AppendLine("            Trabajadores.Status                                     AS Status, ")
            loConsulta.AppendLine("            Trabajadores.Cedula                                     AS Cedula, ")
            loConsulta.AppendLine("            Trabajadores.Fec_Ini                                    AS Fec_Ini, ")
            loConsulta.AppendLine("            CASE WHEN Trabajadores.Status = 'L'  ")
            loConsulta.AppendLine("                THEN COALESCE(Liquidaciones.Fecha, Trabajadores.Fec_Fin) ")
            loConsulta.AppendLine("                ELSE @ldFechaFin ")
            loConsulta.AppendLine("            END                                                     AS Fec_Fin, ")
            loConsulta.AppendLine("            COALESCE(Sueldo_Mensual.Val_Num, 0)                     AS Ultimo_Sueldo_Mensual, ")
            loConsulta.AppendLine("            ROUND(COALESCE(Sueldo_Mensual.Val_Num, 0)/30.0, 2)      AS Ultimo_Sueldo_Diario, ")
            loConsulta.AppendLine("            COALESCE(I_Prestaciones.Val_Num, 0)                     AS Inicial_Prestaciones, ")
            loConsulta.AppendLine("            COALESCE(I_Anticipo_Prestaciones.Val_Num, 0)            AS Inicial_Anticipo_Prestaciones, ")
            loConsulta.AppendLine("            COALESCE(I_Intereses_Prestaciones.Val_Num, 0)           AS Inicial_Intereses_Prestaciones, ")
            loConsulta.AppendLine("            COALESCE(I_Dias_Prestaciones.Val_Num, 0)                AS Inicial_Dias_Prestaciones, ")
            loConsulta.AppendLine("            (CASE WHEN (@llAumentarPorcentaje = 1)")
            loConsulta.AppendLine("                THEN @lnAumentoPorcentaje")
            loConsulta.AppendLine("                ELSE (CASE WHEN COALESCE(Sueldo_Mensual.Val_Num, 0)>0 ")
            loConsulta.AppendLine("                    THEN ROUND(@lnAumentoMonto*100/COALESCE(Sueldo_Mensual.Val_Num, 0), 2)")
            loConsulta.AppendLine("                    ELSE 0")
            loConsulta.AppendLine("                END)")
            loConsulta.AppendLine("            END)                                                    AS Porcentaje_Aumento, ")
            loConsulta.AppendLine("            (CASE WHEN (@llAumentarPorcentaje = 1)")
            loConsulta.AppendLine("                THEN ROUND(COALESCE(Sueldo_Mensual.Val_Num, 0)*@lnAumentoPorcentaje/100, 2)")
            loConsulta.AppendLine("                ELSE @lnAumentoMonto")
            loConsulta.AppendLine("            END)                                                    AS Monto_Aumento, ")
            loConsulta.AppendLine("            (CASE WHEN (@llAumentarPorcentaje = 1)")
            loConsulta.AppendLine("                THEN ROUND(COALESCE(Sueldo_Mensual.Val_Num, 0)*(100+@lnAumentoPorcentaje)/100, 2)")
            loConsulta.AppendLine("                ELSE ROUND(COALESCE(Sueldo_Mensual.Val_Num, 0)+@lnAumentoMonto, 2) ")
            loConsulta.AppendLine("            END)                                                    AS Nuevo_Sueldo ")
            loConsulta.AppendLine("FROM        Trabajadores ")
            loConsulta.AppendLine("    LEFT JOIN Liquidaciones ")
            loConsulta.AppendLine("        ON  Liquidaciones.Cod_Tra = Trabajadores.Cod_Tra ")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina I_Prestaciones ")
            loConsulta.AppendLine("        ON  I_Prestaciones.Cod_Tra = Trabajadores.Cod_Tra ")
            loConsulta.AppendLine("        AND I_Prestaciones.Cod_Cam = 'Z002' ")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina I_Anticipo_Prestaciones ")
            loConsulta.AppendLine("        ON  I_Anticipo_Prestaciones.Cod_Tra = Trabajadores.Cod_Tra ")
            loConsulta.AppendLine("        AND I_Anticipo_Prestaciones.Cod_Cam = 'Z004' ")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina I_Intereses_Prestaciones ")
            loConsulta.AppendLine("        ON  I_Intereses_Prestaciones.Cod_Tra = Trabajadores.Cod_Tra ")
            loConsulta.AppendLine("        AND I_Intereses_Prestaciones.Cod_Cam = 'Z003' ")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina I_Dias_Prestaciones ")
            loConsulta.AppendLine("        ON  I_Dias_Prestaciones.Cod_Tra = Trabajadores.Cod_Tra ")
            loConsulta.AppendLine("        AND I_Dias_Prestaciones.Cod_Cam = 'Z012' ")
            loConsulta.AppendLine("    LEFT JOIN Renglones_Campos_Nomina Sueldo_Mensual ")
            loConsulta.AppendLine("        ON  Sueldo_Mensual.Cod_Tra = Trabajadores.Cod_Tra ")
            loConsulta.AppendLine("        AND Sueldo_Mensual.Cod_Cam = 'A001' ")
            loConsulta.AppendLine("WHERE   Trabajadores.Cod_Tra BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("    AND Trabajadores.Status IN ( " & lcParametro1Desde & " )")
            loConsulta.AppendLine("    AND Trabajadores.Tip_Tra = 'Trabajador'")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Calcula Los meses y años de antiguedad usados para prestaciones a último/nuevo sueldo ")
            loConsulta.AppendLine("-- y los días de prestaciones a último/nuevo Sueldo  ")
            loConsulta.AppendLine("UPDATE  #tmpTrabajadores ")
            loConsulta.AppendLine("SET     Dias_Prestaciones_Mes = COALESCE((SELECT TOP 1 Val_Num FROM Constantes_Locales WHERE cod_con = 'C008'), 5), ")
            loConsulta.AppendLine("        Dias_Prestaciones_Ano = COALESCE((SELECT TOP 1 Val_Num FROM Constantes_Locales WHERE cod_con = 'C005'), 30); ")
            loConsulta.AppendLine(" ")
            loConsulta.AppendLine("UPDATE  #tmpTrabajadores  ")
            loConsulta.AppendLine("SET     Antiguedad = DATEDIFF(MONTH, DATEADD(DAY, -DAY(Ingreso)+1, Ingreso ), DATEADD(DAY, -DAY(Ingreso)+1+1, Egreso ) );")
            loConsulta.AppendLine("         ")
            loConsulta.AppendLine("UPDATE  #tmpTrabajadores  ")
            loConsulta.AppendLine("SET     Meses_Antiguedad = Antiguedad % 12, ")
            loConsulta.AppendLine("        Anos_Antiguedad =  FLOOR( Antiguedad /12.0), ")
            loConsulta.AppendLine("        Anos_Prestaciones = FLOOR( Antiguedad /12.0) + (CASE WHEN (Antiguedad % 12>=6) THEN 1 ELSE 0 END); ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Calcula las prestaciones a último/nuevo sueldo ")
            loConsulta.AppendLine("UPDATE  #tmpTrabajadores  ")
            loConsulta.AppendLine("SET     Prestaciones_Ultimo_Sueldo = ROUND(Anos_Prestaciones*Ultimo_Sueldo_Mensual, 2),")
            loConsulta.AppendLine("        Prestaciones_Nuevo_Sueldo =  ROUND(Anos_Prestaciones*Nuevo_Sueldo, 2),")
            loConsulta.AppendLine("        Diferencia_Pre_Nuevo_Sueldo = ROUND(Anos_Prestaciones*Nuevo_Sueldo, 2) ")
            loConsulta.AppendLine("                                    - ROUND(Anos_Prestaciones*Ultimo_Sueldo_Mensual, 2)")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Seleccion Final")
            loConsulta.AppendLine("SELECT  #tmpTrabajadores.Cod_Tra                                    AS Cod_Tra,")
            loConsulta.AppendLine("        #tmpTrabajadores.Nom_Tra                                    AS Nom_Tra,")
            loConsulta.AppendLine("        #tmpTrabajadores.Status                                     AS Status,")
            loConsulta.AppendLine("        #tmpTrabajadores.Ultimo_Sueldo_Mensual                      AS Ultimo_Sueldo_Mensual,")
            loConsulta.AppendLine("        (   #tmpTrabajadores.Inicial_Prestaciones ")
            loConsulta.AppendLine("          - #tmpTrabajadores.Inicial_Anticipo_Prestaciones")
            loConsulta.AppendLine("          + COALESCE(Movimientos.Abono_Prestaciones, 0)")
            loConsulta.AppendLine("          - COALESCE(Movimientos.Anticipo_Prestaciones, 0))         AS Prestaciones_Fondo,")
            loConsulta.AppendLine("        #tmpTrabajadores.Prestaciones_Ultimo_Sueldo                 AS Prestaciones_Ultimo_Sueldo,")
            loConsulta.AppendLine("        #tmpTrabajadores.Porcentaje_Aumento                         AS Porcentaje_Aumento,")
            loConsulta.AppendLine("        #tmpTrabajadores.Monto_Aumento                              AS Monto_Aumento,")
            loConsulta.AppendLine("        #tmpTrabajadores.Nuevo_Sueldo                               AS Nuevo_Sueldo,")
            loConsulta.AppendLine("        #tmpTrabajadores.Prestaciones_Nuevo_Sueldo                  AS Prestaciones_Nuevo_Sueldo,")
            loConsulta.AppendLine("        #tmpTrabajadores.Diferencia_Pre_Nuevo_Sueldo                AS Diferencia_Pre_Nuevo_Sueldo,")
            loConsulta.AppendLine("        (   #tmpTrabajadores.Inicial_Intereses_Prestaciones ")
            loConsulta.AppendLine("          + COALESCE(Movimientos.Interes_Prestaciones,2) ")
            loConsulta.AppendLine("          - COALESCE(Movimientos.Interes_Pagado,2))                 AS Saldo_Intereses,")
            loConsulta.AppendLine("        (   #tmpTrabajadores.Prestaciones_Nuevo_Sueldo ")
            loConsulta.AppendLine("          + #tmpTrabajadores.Inicial_Intereses_Prestaciones ")
            loConsulta.AppendLine("          + COALESCE(Movimientos.Interes_Prestaciones,2) ")
            loConsulta.AppendLine("          - COALESCE(Movimientos.Interes_Pagado,2))                 AS Total_Deuda")
            loConsulta.AppendLine("FROM    #tmpTrabajadores")
            loConsulta.AppendLine("    LEFT JOIN (SELECT  Recibos.Cod_Tra                                  AS Cod_Tra, ")
            loConsulta.AppendLine("                   SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'A450') ")
            loConsulta.AppendLine("                       THEN Renglones_Recibos.Mon_Net ")
            loConsulta.AppendLine("                       ELSE 0 ")
            loConsulta.AppendLine("                   END)                                                 AS Abono_Prestaciones, ")
            loConsulta.AppendLine("                   SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'B009') ")
            loConsulta.AppendLine("                       THEN Renglones_Recibos.Mon_Net ")
            loConsulta.AppendLine("                       ELSE 0 ")
            loConsulta.AppendLine("                   END)                                                 AS Anticipo_Prestaciones, ")
            loConsulta.AppendLine("                   SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'A010') ")
            loConsulta.AppendLine("                       THEN Renglones_Recibos.Mon_Net ")
            loConsulta.AppendLine("                       ELSE 0 ")
            loConsulta.AppendLine("                   END)                                                 AS Interes_Prestaciones, ")
            loConsulta.AppendLine("                   SUM(CASE WHEN (Renglones_Recibos.Cod_Con = 'A100') ")
            loConsulta.AppendLine("                       THEN Renglones_Recibos.Mon_Net ")
            loConsulta.AppendLine("                       ELSE 0 ")
            loConsulta.AppendLine("                   END)                                                 AS Interes_Pagado ")
            loConsulta.AppendLine("            FROM    Recibos ")
            loConsulta.AppendLine("                JOIN #tmpTrabajadores ON #tmpTrabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loConsulta.AppendLine("                JOIN Renglones_Recibos ON Renglones_Recibos.Documento = Recibos.Documento ")
            loConsulta.AppendLine("            WHERE   Recibos.Status = 'Confirmado' ")
            loConsulta.AppendLine("                AND Recibos.Fecha <= @ldFechaFin")
            loConsulta.AppendLine("            GROUP BY Recibos.Cod_Tra) AS Movimientos")
            loConsulta.AppendLine("        ON Movimientos.Cod_Tra = #tmpTrabajadores.Cod_Tra")
            loConsulta.AppendLine("ORDER BY #tmpTrabajadores.Cod_Tra ASC")
            loConsulta.AppendLine("")
        
            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString, "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rEstimaciones_Sueldos", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrEstimaciones_Sueldos.ReportSource = loObjetoReporte

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
' RJG: 24/10/14: Codigo inicial.                                                            '
'-------------------------------------------------------------------------------------------'
