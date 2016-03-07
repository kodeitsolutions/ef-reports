'-------------------------------------------------------------------------------------------'
' Inicio del codigo
'-------------------------------------------------------------------------------------------'
' Importando librerias 
'-------------------------------------------------------------------------------------------'
Imports System.Data

'-------------------------------------------------------------------------------------------'
' Inicio de clase "rPromocionesMarketing"
'-------------------------------------------------------------------------------------------'
Partial Class rPromocionesMarketing
    Inherits vis2Formularios.frmReporte

    Dim loObjetoReporte As CrystalDecisions.CrystalReports.Engine.ReportDocument

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim loConsulta As New StringBuilder()
        Dim lcParametro0Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(0))
        Dim lcParametro0Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(0))
        Dim lcParametro1Desde As String = goServicios.mObtenerListaFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(1))
        Dim lcParametro2Desde As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosIniciales(2))
        Dim lcParametro2Hasta As String = goServicios.mObtenerCampoFormatoSQL(cusAplicacion.goReportes.paParametrosFinales(2))

        Dim lcOrdenamiento As String = cusAplicacion.goReportes.pcOrden

        Try
            loConsulta.AppendLine("SELECT   Promociones.Documento AS Documento,")
            loConsulta.AppendLine("         Promociones.Fec_Ini,")
            loConsulta.AppendLine("         Promociones.Fec_Fin,")
            loConsulta.AppendLine("         Promociones.Status AS Status,")
            loConsulta.AppendLine("         Promociones.Tipo,")
            loConsulta.AppendLine("         Promociones.Clase")
            loConsulta.AppendLine("FROM	    Promociones")
            loConsulta.AppendLine("WHERE      Promociones.Documento BETWEEN " & lcParametro0Desde & " AND " & lcParametro0Hasta)
            loConsulta.AppendLine("     AND      Promociones.Status IN (" & lcParametro1Desde & ")")
            loConsulta.AppendLine("     AND      Promociones.Fec_Ini BETWEEN " & lcParametro2Desde & " AND " & lcParametro2Hasta)
            loConsulta.AppendLine("ORDER BY " & lcOrdenamiento)
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")
            loConsulta.AppendLine("")

            'Me.mEscribirConsulta(loConsulta.ToString())

            Dim loServicios As New cusDatos.goDatos

            Dim laDatosReporte As DataSet = loServicios.mObtenerTodosSinEsquema(loConsulta.ToString(), "curReportes")

            loObjetoReporte = cusAplicacion.goReportes.mCargarReporte("rPromocionesMarketing", laDatosReporte)

            Me.mTraducirReporte(loObjetoReporte)

            Me.mFormatearCamposReporte(loObjetoReporte)

            Me.crvrPromocionesMarketing.ReportSource = loObjetoReporte

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

