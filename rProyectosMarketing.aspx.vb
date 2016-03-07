'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rProyectosMarketing"
'-------------------------------------------------------------------------------------------'
Partial Class rProyectosMarketing
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()

        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))

        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))


        Dim lcParametro3Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(3))
        Dim lcParametro3Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(3))
        Dim lcParametro4Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(4))


        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try
            loConsulta.AppendLine("SELECT   Proyectos.Cod_Pro AS Codigo,")
            loConsulta.AppendLine("         Proyectos.Nom_Pro AS Nombre,")
            loConsulta.AppendLine("         Proyectos.Responsable AS Responsable,")
            loConsulta.AppendLine("         Proyectos.Fec_Ini,")
            loConsulta.AppendLine("         Proyectos.Fec_Fin,")
            loConsulta.AppendLine("         Proyectos.Por_Eje,")
            loConsulta.AppendLine("         Proyectos.Status AS Status")
            loConsulta.AppendLine("FROM	    Proyectos")
            loConsulta.AppendLine("WHERE      Proyectos.Cod_Pro BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("     AND      Proyectos.Status IN (" & lcParametro1Desde & " )")
            loConsulta.AppendLine("     AND      Proyectos.Etapa IN (" & lcParametro2Desde & " )")
            loConsulta.AppendLine("     AND      Proyectos.Fec_Ini BETWEEN " & lcParametro3Desde & " AND " & lcParametro3Hasta)
            loConsulta.AppendLine("     AND      Proyectos.Prioridad IN (" & lcParametro4Desde & ")")
            loConsulta.AppendLine("    AND Proyectos.Adicional = ''    ")
            'loConsulta.AppendLine("     AND      Proyectos.Adicional IN ('','Marketing')")
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rProyectosMarketing", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrProyectosMarketing.ReportSource = loObjetoReporte

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
' JAC : 29/07/15 : Codigo inicial
'-------------------------------------------------------------------------------------------'
' EAG: 11/09/15: Se agregó filtro Proyectos.Adicional = '' .                                '
'-------------------------------------------------------------------------------------------'

