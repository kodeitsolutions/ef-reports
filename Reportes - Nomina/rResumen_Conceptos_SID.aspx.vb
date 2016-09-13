'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rResumen_Conceptos_SID"
'-------------------------------------------------------------------------------------------'
Partial Class rResumen_Conceptos_SID
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try

            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
            'Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))
            Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
            Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
            Dim lcParametro4Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))
            Dim lcParametro4Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(4))
            Dim lcParametro5Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(5))
            Dim lcParametro5Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(5))
            Dim lcParametro6Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(6))
            Dim lcParametro6Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(6))
            Dim lcParametro7Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(7))
            Dim lcParametro7Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(7))
            Dim lcParametro8Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loComandoSeleccionar As New StringBuilder()

            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("CREATE TABLE #tmpResumen(   Cod_Con CHAR(10) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Nom_Con CHAR(100) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Tipo CHAR(30) COLLATE DATABASE_DEFAULT,")
            loComandoSeleccionar.AppendLine("                            Asignacion DECIMAL(28, 10),")
            loComandoSeleccionar.AppendLine("                            Deduccion DECIMAL(28, 10),")
            loComandoSeleccionar.AppendLine("                            Otro DECIMAL(28, 10),")
            loComandoSeleccionar.AppendLine("                            Saldo DECIMAL(28, 10),")
            loComandoSeleccionar.AppendLine("                            Cod_Tra CHAR(10) COLLATE DATABASE_DEFAULT);")
            loComandoSeleccionar.AppendLine("                            ")
            loComandoSeleccionar.AppendLine("INSERT INTO #tmpResumen(Cod_Con, Nom_Con, Tipo, Asignacion, Deduccion, Otro, Saldo, Cod_Tra)")
            loComandoSeleccionar.AppendLine("SELECT      Conceptos_Nomina.Cod_Con                                AS Cod_Con,")
            loComandoSeleccionar.AppendLine("            Conceptos_Nomina.Nom_Con                                AS Nom_Con,")
            loComandoSeleccionar.AppendLine("            Conceptos_Nomina.tipo                                   AS tipo,")
            loComandoSeleccionar.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loComandoSeleccionar.AppendLine("                WHEN 'Asignacion' THEN Renglones_Recibos.mon_net ")
            loComandoSeleccionar.AppendLine("                ELSE 0 END)                                         AS Asignacion,")
            loComandoSeleccionar.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loComandoSeleccionar.AppendLine("                WHEN 'Deduccion' THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("                WHEN 'Retencion' THEN Renglones_Recibos.mon_net")
            loComandoSeleccionar.AppendLine("                ELSE 0 END)                                         AS Deduccion,")
            loComandoSeleccionar.AppendLine("            SUM(CASE Conceptos_Nomina.tipo ")
            loComandoSeleccionar.AppendLine("                WHEN 'Otro' THEN Renglones_Recibos.mon_net ")
            loComandoSeleccionar.AppendLine("                ELSE 0 END)                                         AS Otro,")
            loComandoSeleccionar.AppendLine("            COALESCE(Prestamos.Mon_sal, 0)                          AS Saldo, ")
            loComandoSeleccionar.AppendLine("            Recibos.cod_tra                                         AS Cod_Tra")
            loComandoSeleccionar.AppendLine("FROM        Renglones_Recibos")
            loComandoSeleccionar.AppendLine("    JOIN    Recibos ")
            loComandoSeleccionar.AppendLine("        ON  Recibos.Documento = Renglones_Recibos.Documento")
            loComandoSeleccionar.AppendLine("    JOIN    Conceptos_Nomina ")
            loComandoSeleccionar.AppendLine("        ON  Conceptos_Nomina.cod_con = Renglones_Recibos.cod_con ")
            loComandoSeleccionar.AppendLine("    JOIN    Trabajadores ")
            loComandoSeleccionar.AppendLine("        ON  Trabajadores.cod_tra = Recibos.cod_tra ")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Detalles_Prestamos ")
            loComandoSeleccionar.AppendLine("        ON  Detalles_Prestamos.documento = Renglones_Recibos.doc_ori")
            loComandoSeleccionar.AppendLine("        AND Detalles_Prestamos.renglon = Renglones_Recibos.ren_ori")
            loComandoSeleccionar.AppendLine("        AND Renglones_Recibos.Tip_Ori = 'Prestamos'")
            loComandoSeleccionar.AppendLine("    LEFT JOIN Prestamos ")
            loComandoSeleccionar.AppendLine("        ON  Detalles_Prestamos.documento = Prestamos.documento")
            loComandoSeleccionar.AppendLine("        AND Prestamos.status = 'Confirmado'")
            loComandoSeleccionar.AppendLine("        AND Prestamos.Mon_Sal > 0")
            loComandoSeleccionar.AppendLine("WHERE		 Conceptos_Nomina.Cod_Con BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loComandoSeleccionar.AppendLine("        AND Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loComandoSeleccionar.AppendLine("        AND Conceptos_Nomina.Tipo IN (" & lcParametro2Desde & ")")
            loComandoSeleccionar.AppendLine("        AND Recibos.Cod_Con BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loComandoSeleccionar.AppendLine("        AND Trabajadores.Cod_Dep BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loComandoSeleccionar.AppendLine("        AND Trabajadores.Cod_Suc BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loComandoSeleccionar.AppendLine("        AND Trabajadores.Cod_Tra BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loComandoSeleccionar.AppendLine("        AND Conceptos_Nomina.Tipo BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loComandoSeleccionar.AppendLine("        AND Recibos.Status IN (" & lcParametro8Desde & ")")
            loComandoSeleccionar.AppendLine("GROUP BY    conceptos_Nomina.Cod_Con,")
            loComandoSeleccionar.AppendLine("            conceptos_Nomina.Nom_Con,")
            loComandoSeleccionar.AppendLine("            conceptos_Nomina.Tipo, ")
            loComandoSeleccionar.AppendLine("            COALESCE(Prestamos.Documento, ''),")
            loComandoSeleccionar.AppendLine("            COALESCE(Prestamos.Mon_sal, 0),")
            loComandoSeleccionar.AppendLine("            Recibos.cod_tra")
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("SELECT  Cod_Con, Nom_Con, Tipo, ")
            loComandoSeleccionar.AppendLine("        SUM(Asignacion) AS Asignacion, ")
            loComandoSeleccionar.AppendLine("        SUM(Deduccion)  AS Deduccion, ")
            loComandoSeleccionar.AppendLine("        SUM(Otro)       AS Otro, ")
            loComandoSeleccionar.AppendLine("        SUM(Saldo)      AS Saldo, ")
            loComandoSeleccionar.AppendLine("        COUNT(DISTINCT cod_tra) AS Trabajadores,")
            loComandoSeleccionar.AppendLine("        (SELECT COUNT(DISTINCT Cod_Tra) FROM #tmpResumen) AS Integrantes,")
            loComandoSeleccionar.AppendLine("        CAST(" & lcParametro1Desde & " AS DATETIME) AS Desde, ")
            loComandoSeleccionar.AppendLine("        CAST(" & lcParametro1Hasta & " AS DATETIME) AS Hasta ")
            loComandoSeleccionar.AppendLine("FROM    #tmpResumen")
            loComandoSeleccionar.AppendLine("GROUP BY Cod_Con, Nom_Con, Tipo")
            loComandoSeleccionar.AppendLine("ORDER BY(CASE tipo")
            loComandoSeleccionar.AppendLine("            WHEN 'Asignacion' THEN 0")
            loComandoSeleccionar.AppendLine("            WHEN 'Deduccion' THEN 1")
            loComandoSeleccionar.AppendLine("            WHEN 'Retencion' THEN 1")
            loComandoSeleccionar.AppendLine("            ELSE 2 END) ASC, ")
            loComandoSeleccionar.AppendLine("        cod_con ASC")
            loComandoSeleccionar.AppendLine("")
            'loComandoSeleccionar.AppendLine("ORDER BY	" & lcOrdenamiento)
            loComandoSeleccionar.AppendLine("")
            loComandoSeleccionar.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loComandoSeleccionar.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rResumen_Conceptos_SID", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrResumen_Conceptos_SID.ReportSource = loObjetoReporte

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
' RJG: 13/06/13: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
