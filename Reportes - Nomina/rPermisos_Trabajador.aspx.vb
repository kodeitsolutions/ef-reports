'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPermisos_Trabajador"
'-------------------------------------------------------------------------------------------'
Partial Class rPermisos_Trabajador
    Inherits vis2formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
            Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
            Dim lcParametro1Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1), goServicios.enuOpcionesRedondeo.KN_FechaInicioDelDia)
            Dim lcParametro1Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(1), goServicios.enuOpcionesRedondeo.KN_FechaFinDelDia)
            Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
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

            Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden
            Dim loConsulta As New StringBuilder()

            loConsulta.AppendLine("SELECT      Permisos.Documento    AS Documento,	")
            loConsulta.AppendLine("            Permisos.fecha        AS Fecha,		")
            loConsulta.AppendLine("            Permisos.[Status]     AS Estatus,		")
            loConsulta.AppendLine("            Permisos.Cod_Tra      AS Cod_Tra,		")
            loConsulta.AppendLine("            Trabajadores.Nom_Tra AS Nom_Tra,		")
            loConsulta.AppendLine("            Permisos.Fec_Ini      AS Fec_Ini,		")
            loConsulta.AppendLine("            Permisos.Fec_Fin      AS Fec_Fin,		")
            loConsulta.AppendLine("            Permisos.Dias         AS Dias,		")
            loConsulta.AppendLine("            Permisos.Cod_Rev      AS Cod_Rev,		")
            loConsulta.AppendLine("            Permisos.Motivo       AS Motivo,		")
            loConsulta.AppendLine("            Motivos.Nom_Mot      AS Nom_Mot,		")
            loConsulta.AppendLine("            Permisos.Comentario   AS Comentario	")
            loConsulta.AppendLine("FROM	       Permisos ")
            loConsulta.AppendLine("    JOIN    Trabajadores ")
            loConsulta.AppendLine("     ON     Trabajadores.Cod_Tra = Permisos.Cod_Tra")
            loConsulta.AppendLine("    JOIN    Motivos ")
            loConsulta.AppendLine("        ON  Motivos.Cod_Mot = Permisos.Motivo")
            loConsulta.AppendLine("WHERE       Permisos.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("        AND Permisos.Fecha BETWEEN " & lcParametro1Desde & " AND " & lcParametro1Hasta)
            loConsulta.AppendLine("        AND Permisos.Status IN (" & lcParametro2Desde & ")")
            loConsulta.AppendLine("        AND Permisos.Motivo BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("        AND Permisos.Cod_Tra BETWEEN " & lcParametro4Desde & " AND " & lcParametro4Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Con BETWEEN " & lcParametro5Desde & " AND " & lcParametro5Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Dep BETWEEN " & lcParametro6Desde & " AND " & lcParametro6Hasta)
            loConsulta.AppendLine("        AND Trabajadores.Cod_Car BETWEEN " & lcParametro7Desde & " AND " & lcParametro7Hasta)
            loConsulta.AppendLine("        AND Permisos.Cod_Suc BETWEEN " & lcParametro8Desde & " AND " & lcParametro8Hasta)
            loConsulta.AppendLine("        AND Permisos.Cod_Rev BETWEEN " & lcParametro9Desde & " AND " & lcParametro9Hasta)
            loConsulta.AppendLine("ORDER BY	   Permisos.Cod_Tra ASC, " & lcOrdenamiento)

            Dim loServicios As New cusDatos.goDatos

            'Me.mEscribirConsulta(loComandoSeleccionar.ToString())
            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")


            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPermisos_Trabajador", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPermisos_Trabajador.ReportSource = loObjetoReporte

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
' RJG: 14/08/14: Codigo inicial																'
'-------------------------------------------------------------------------------------------'
