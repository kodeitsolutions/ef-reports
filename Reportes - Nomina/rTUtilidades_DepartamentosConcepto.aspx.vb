'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rTUtilidades_DepartamentosConcepto"
'-------------------------------------------------------------------------------------------'
Partial Class rTUtilidades_DepartamentosConcepto
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
            Dim lcParametro8Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(8))
            Dim lcParametro8Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(8))
            Dim lcParametro9Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(9))
            Dim lcParametro9Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(9))
            Dim lcParametro10Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(10))
            Dim lcParametro10Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(10))
            Dim lcParametro11Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(11))
            Dim lcParametro11Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(11))

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("")
            loConsulta.AppendLine("SELECT      Conceptos_Nomina.Cod_Con                AS Cod_Con,")
            loConsulta.AppendLine("            Conceptos_Nomina.Nom_Con                AS Nom_Con,")
            loConsulta.AppendLine("            Conceptos_Nomina.tipo                   AS tipo,")
            loConsulta.AppendLine("            Departamentos_Nomina.Cod_Dep            AS Cod_Dep,")
            loConsulta.AppendLine("            Departamentos_Nomina.Nom_Dep            AS Nom_Dep,")
            loConsulta.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loConsulta.AppendLine("                WHEN 'Asignacion'")
            loConsulta.AppendLine("                THEN Renglones_Recibos.mon_net")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)                                   AS Asignacion,")
            loConsulta.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loConsulta.AppendLine("                WHEN 'Deduccion'")
            loConsulta.AppendLine("                THEN - Renglones_Recibos.mon_net")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)                                   AS Deduccion,")
            loConsulta.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loConsulta.AppendLine("                WHEN 'Retencion'")
            loConsulta.AppendLine("                THEN - Renglones_Recibos.mon_net")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)                                   AS Retencion,")
            loConsulta.AppendLine("            SUM(CASE Conceptos_Nomina.tipo")
            loConsulta.AppendLine("                WHEN 'Otro'")
            loConsulta.AppendLine("                THEN Renglones_Recibos.mon_net")
            loConsulta.AppendLine("                ELSE 0")
            loConsulta.AppendLine("            END)                                   AS Otro")
            loConsulta.AppendLine("FROM        Renglones_Recibos")
            loConsulta.AppendLine("    JOIN    Recibos ")
            loConsulta.AppendLine("        ON  Recibos.Documento = Renglones_Recibos.Documento")
            loConsulta.AppendLine("    JOIN    Conceptos_Nomina ")
            loConsulta.AppendLine("        ON  Conceptos_Nomina.cod_con = Renglones_Recibos.cod_con ")
            loConsulta.AppendLine("    JOIN    Trabajadores ")
            loConsulta.AppendLine("        ON  Trabajadores.cod_tra = Recibos.cod_tra ")
            loConsulta.AppendLine("    JOIN    Departamentos_Nomina ")
            loConsulta.AppendLine("        ON  Departamentos_Nomina.cod_dep = Trabajadores.cod_dep ")
            'loConsulta.AppendLine("    JOIN    Contratos ")
            'loConsulta.AppendLine("        ON  Contratos.Cod_Con = Recibos.Cod_Con ")
            'loConsulta.AppendLine("        AND Contratos.Gen_Lib = 0")
            loConsulta.AppendLine("WHERE	   Conceptos_Nomina.Cod_Con BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Recibos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loConsulta.AppendLine("        AND Conceptos_Nomina.Tipo IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine("        AND (   Conceptos_Nomina.Tipo IN ('Asignacion','Deduccion','Retencion')")
            loConsulta.AppendLine("             OR Conceptos_Nomina.Cod_Con IN ('U001','U002','U003','U004','U204','U301','U302','U303','U304','U403','U404')")
            loConsulta.AppendLine("            )")
            loConsulta.AppendLine("        AND Recibos.Cod_Con BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("        AND Recibos.Cod_Con IN ('80','81','90') ")
            loConsulta.AppendLine("        AND Departamentos_Nomina.Cod_Dep BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Car BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Tra BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loConsulta.AppendLine("        AND Conceptos_Nomina.Cod_Cla BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Suc BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loConsulta.AppendLine("        AND Recibos.Cod_Mon BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta)
            loConsulta.AppendLine("        AND Recibos.Cod_Rev BETWEEN " & lcParametro10Desde & " AND " & lcParametro10Hasta)
            loConsulta.AppendLine("        AND Recibos.Status IN ('Confirmado', 'Procesado')")
            loConsulta.AppendLine("        AND Trabajadores.Cod_Con BETWEEN " & lcParametro11Desde & " AND " & lcParametro11Hasta)
            loConsulta.AppendLine("GROUP BY    conceptos_Nomina.Cod_Con,")
            loConsulta.AppendLine("            conceptos_Nomina.Nom_Con,")
            loConsulta.AppendLine("            conceptos_Nomina.Tipo,")
            loConsulta.AppendLine("            Departamentos_Nomina.Cod_Dep,")
            loConsulta.AppendLine("            Departamentos_Nomina.Nom_Dep")
            loConsulta.AppendLine("ORDER BY    Departamentos_Nomina.Cod_Dep, ")
            loConsulta.AppendLine("            (CASE Conceptos_Nomina.tipo")
            loConsulta.AppendLine("               WHEN 'Asignacion' THEN 0")
            loConsulta.AppendLine("                WHEN 'Deduccion' THEN 1")
            loConsulta.AppendLine("                WHEN 'Retencion' THEN 1")
            loConsulta.AppendLine("                ELSE 2")
            loConsulta.AppendLine("             END) ASC,")
            loConsulta.AppendLine("             cod_con ASC")
            loConsulta.AppendLine("             ")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loConsulta.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rTUtilidades_DepartamentosConcepto", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrTUtilidades_DepartamentosConcepto.ReportSource = loObjetoReporte

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
' RJG: 06/07/15: Codigo inicial, a partir de rTotales_TrabajadoresConcepto.					'
'-------------------------------------------------------------------------------------------'
