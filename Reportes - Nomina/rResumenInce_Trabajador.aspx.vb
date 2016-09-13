﻿'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rResumenInce_Trabajador"
'-------------------------------------------------------------------------------------------'
Partial Class rResumenInce_Trabajador
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()
            
            Dim ldFecha As Date = CDate(cusAplicacion.goReportes.paParametrosIniciales(1))
            Dim lcPeriodo AS String
            If (ldFecha.Month >= 1 And ldFecha.Month <= 3) Then
                lcPeriodo = "Enero/[AÑO], Febrero/[AÑO], Marzo/[AÑO]"
            ElseIf (ldFecha.Month >= 4 And ldFecha.Month <= 6) Then
                lcPeriodo = "Abril/[AÑO], Mayo/[AÑO], Junio/[AÑO]" 
            ElseIf (ldFecha.Month >= 7 And ldFecha.Month <= 9) Then
                lcPeriodo = "Julio/[AÑO], Agosto/[AÑO], Septiembre/[AÑO]"
            Else
                lcPeriodo = "Octubre/[AÑO], Noviembre/[AÑO], Diciembre/[AÑO]" 
            End If
            
            lcPeriodo = lcPeriodo.Replace("[AÑO]", ldFecha.Year.ToString("0000"))

            loConsulta.AppendLine("")
            loConsulta.AppendLine("DECLARE @lnPorcentajeTrabajador DECIMAL(28,10);")
            loConsulta.AppendLine("DECLARE @lnPorcentajePatrono DECIMAL(28,10);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SET @lnPorcentajeTrabajador = (SELECT TOP 1 Val_Num FROM Constantes_Locales WHERE Cod_Con = 'R004');")
            loConsulta.AppendLine("SET @lnPorcentajePatrono = (SELECT TOP 1 Val_Num FROM Constantes_Locales WHERE Cod_Con = 'U004');")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("CREATE TABLE #tmpConceptos (Cod_Tra CHAR(10)    COLLATE DATABASE_DEFAULT NOT NULL,")
            loConsulta.AppendLine("                            Nom_Tra CHAR(100)   COLLATE DATABASE_DEFAULT NOT NULL,")
            loConsulta.AppendLine("                            Cedula CHAR(30)     COLLATE DATABASE_DEFAULT NOT NULL,")
            loConsulta.AppendLine("                            Contrato CHAR(10)   COLLATE DATABASE_DEFAULT NOT NULL,")
            loConsulta.AppendLine("                            Cod_Con CHAR(10)    COLLATE DATABASE_DEFAULT NOT NULL,")
            loConsulta.AppendLine("                            Utilidades BIT,")
            loConsulta.AppendLine("                            Mon_Net DECIMAL (28,10));")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("INSERT INTO #tmpConceptos(Cod_Tra, Nom_Tra, Cedula, Contrato, Cod_Con, Utilidades, Mon_Net)")
            loConsulta.AppendLine("SELECT      Trabajadores.Cod_Tra                AS Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra                AS Nom_Tra,")
            loConsulta.AppendLine("            Trabajadores.Cedula                 AS Cedula,")
            loConsulta.AppendLine("            Recibos.Cod_Con                     AS Contrato,")
            loConsulta.AppendLine("            Conceptos_Nomina.Cod_Con            AS Cod_Con,")
            loConsulta.AppendLine("            Conceptos_Nomina.Utilidades         AS Utilidades,")
            loConsulta.AppendLine("            SUM(Renglones_Recibos.Mon_Net)      AS Mon_Net")
            loConsulta.AppendLine("FROM        Renglones_Recibos")
            loConsulta.AppendLine("    JOIN    Recibos ")
            loConsulta.AppendLine("        ON  Recibos.Documento = Renglones_Recibos.Documento")
            loConsulta.AppendLine("    JOIN    Conceptos_Nomina")
            loConsulta.AppendLine("        ON  Conceptos_Nomina.Cod_Con = Renglones_Recibos.Cod_Con")
            loConsulta.AppendLine("    JOIN    Trabajadores ")
            loConsulta.AppendLine("        ON  Trabajadores.Cod_Tra = Recibos.Cod_Tra ")
            loConsulta.AppendLine("WHERE       Recibos.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loConsulta.AppendLine("        AND Recibos.Cod_Con BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Tra BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Dep BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Suc BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("        AND Recibos.Cod_Rev BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loConsulta.AppendLine("        AND Recibos.Status IN ('Confirmado', 'Procesado')")
            loConsulta.AppendLine("GROUP BY    Trabajadores.Cod_Tra,")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra,")
            loConsulta.AppendLine("            Trabajadores.Cedula,")
            loConsulta.AppendLine("            Recibos.Cod_Con,")
            loConsulta.AppendLine("            Conceptos_Nomina.Utilidades,")
            loConsulta.AppendLine("            Conceptos_Nomina.Cod_Con")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("-- Para acelerar las siguientes uniones")
            loConsulta.AppendLine("ALTER TABLE #tmpConceptos ADD CONSTRAINT PK_tmpConceptos_Contrato_Cod_Con PRIMARY KEY CLUSTERED (Contrato, Cod_Con, Cod_Tra);")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT Cod_Tra, Nom_Tra, Cedula, Posicion, Grupo, Porcentaje, ")
            loConsulta.AppendLine("       SUM(Base) Base, SUM(Ince) Ince, ")
            loConsulta.AppendLine("       " & goServicios.mObtenerCampoFormatoSQL(lcPeriodo) & " AS Periodo")
            loConsulta.AppendLine("FROM ( ")
            loConsulta.AppendLine("    SELECT      Cod_tra, Nom_tra, Cedula,")
            loConsulta.AppendLine("                1 AS Posicion,")
            loConsulta.AppendLine("                @lnPorcentajeTrabajador AS Porcentaje,")
            loConsulta.AppendLine("                'Utilidades (Retención al trabajador)' AS Grupo,")
            loConsulta.AppendLine("                (CASE ")
            loConsulta.AppendLine("                    WHEN Cod_Con IN ('A200') THEN Mon_Net ")
            loConsulta.AppendLine("                    WHEN Cod_Con IN ('D200') THEN -Mon_Net ")
            loConsulta.AppendLine("                    ELSE 0")
            loConsulta.AppendLine("                END)                        AS Base,")
            loConsulta.AppendLine("                (CASE WHEN Cod_Con IN ('R004')")
            loConsulta.AppendLine("                    THEN Mon_Net ELSE 0")
            loConsulta.AppendLine("                END)                        AS Ince")
            loConsulta.AppendLine("    FROM        #tmpConceptos")
            loConsulta.AppendLine("    WHERE       Contrato IN ('90','96')")
            loConsulta.AppendLine("        AND     Cod_Con IN ('A200', 'D200', 'R004') ")
            loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("    SELECT      Cod_tra, Nom_tra, Cedula,")
            loConsulta.AppendLine("                2 AS Posicion,")
            loConsulta.AppendLine("                @lnPorcentajePatrono AS Porcentaje,")
            loConsulta.AppendLine("                'Nómina (Aporte patronal)' AS Grupo,")
            loConsulta.AppendLine("                (CASE ")
            loConsulta.AppendLine("                    WHEN Utilidades = 1 THEN Mon_Net ")
            loConsulta.AppendLine("                    ELSE 0")
            loConsulta.AppendLine("                END)                        AS Base,")
            loConsulta.AppendLine("                (CASE WHEN Cod_Con IN ('U004','U304')")
            loConsulta.AppendLine("                    THEN Mon_Net ELSE 0")
            loConsulta.AppendLine("                END)                        AS Ince")
            loConsulta.AppendLine("    FROM        #tmpConceptos")
            loConsulta.AppendLine("    WHERE       Contrato IN ('01','02','03','06', '91')")
            loConsulta.AppendLine("        AND     (Utilidades = 1 OR Cod_Con IN ('U004','U304'))")
            loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("    SELECT      Cod_tra, Nom_tra, Cedula,")
            loConsulta.AppendLine("                3 AS Posicion,")
            loConsulta.AppendLine("                @lnPorcentajeTrabajador AS Porcentaje,")
            loConsulta.AppendLine("                'Liquidación (Retención al trabajador)' AS Grupo,")
            loConsulta.AppendLine("                (CASE ")
            loConsulta.AppendLine("                    WHEN Cod_Con IN ('A402') THEN Mon_Net ")
            loConsulta.AppendLine("                    WHEN Cod_Con IN ('D200') THEN -Mon_Net ")
            loConsulta.AppendLine("                    ELSE 0")
            loConsulta.AppendLine("                END)                        AS Base,")
            loConsulta.AppendLine("                (CASE WHEN Cod_Con IN ('R004')")
            loConsulta.AppendLine("                    THEN Mon_Net ELSE 0")
            loConsulta.AppendLine("                END)                        AS Ince")
            loConsulta.AppendLine("    FROM        #tmpConceptos")
            loConsulta.AppendLine("    WHERE       Contrato IN ('92')")
            loConsulta.AppendLine("        AND     (Cod_Con IN ('A402', 'D200', 'R004'))")
            loConsulta.AppendLine("UNION ALL")
            loConsulta.AppendLine("    SELECT      Cod_tra, Nom_tra, Cedula,")
            loConsulta.AppendLine("                4 AS Posicion,")
            loConsulta.AppendLine("                @lnPorcentajePatrono AS Porcentaje,")
            loConsulta.AppendLine("                'Liquidación (Aporte patronal)' AS Grupo,")
            loConsulta.AppendLine("                (CASE ")
            loConsulta.AppendLine("                    WHEN Cod_Con IN ('A407') THEN Mon_Net ")
            loConsulta.AppendLine("                    ELSE 0")
            loConsulta.AppendLine("                END)                        AS Base,")
            loConsulta.AppendLine("                (CASE WHEN Cod_Con IN ('U404')")
            loConsulta.AppendLine("                    THEN Mon_Net ELSE 0")
            loConsulta.AppendLine("                END)                        AS Ince")
            loConsulta.AppendLine("    FROM        #tmpConceptos")
            loConsulta.AppendLine("    WHERE       Contrato IN ('92')")
            loConsulta.AppendLine("        AND     (Cod_Con IN ('A407', 'U404'))")
            loConsulta.AppendLine(") AS Resumen")
            loConsulta.AppendLine("GROUP BY Cod_Tra, Nom_Tra, Cedula, Posicion, Grupo, Porcentaje")
            loConsulta.AppendLine("ORDER BY Posicion, Cod_Tra;")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rResumenInce_Trabajador", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrResumenInce_Trabajador.ReportSource = loObjetoReporte

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
' Fin del codigo																			'
'-------------------------------------------------------------------------------------------'
' RJG: 16/03/15: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
